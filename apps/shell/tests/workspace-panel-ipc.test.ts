import { DOCS_WORKSPACE_AI_ENTRY as ENTRY } from '../../docs/src/shared/workspace-ai-entry'
import { EventEmitter } from 'node:events'
import { mkdtempSync, readFileSync, readdirSync, rmSync, statSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { BrowserWindow, IpcMain, IpcMainInvokeEvent } from 'electron'
import { afterEach, describe, expect, it, vi } from 'vitest'
import {
  ProjectStore,
  type ProjectChatListItem,
  type ProjectListItem,
} from '@genoffice/project-store'
import { defaultAiSettings } from '@genoffice/ai-provider'
import type { AgentStreamCallbacks, AgentStreamRequest } from '@genoffice/agent-core'
import { aiSendGate } from '../../../packages/ui/src/ai-send-gate'
import { WorkspaceConversationHistory } from '../src/main/workspace-ai/conversation-history'
import { WorkspaceExecutionController } from '../src/main/workspace-ai/execution-controller'
import { registerWorkspacePanelIpc } from '../src/main/workspace-ai/panel-ipc'
import { workspacePanelSources } from '../src/main/workspace-ai/panel-sources'
import { workspacePanelBridge } from '../src/preload/workspace-panel-bridge'
import { WORKSPACE_MEMORY_CHANNELS as M } from '../src/shared/workspace-memory-api'
import { workspaceMemoryBridge } from '../src/preload/workspace-memory-bridge'
import { registerWorkspaceMemoryIpc } from '../src/main/workspace-ai/memory-ipc'
import { homeMemoryHost } from '../src/renderer/src/workspace-ai/home-memory-host'
import {
  WorkspacePanelClient,
  connectWorkspacePanel,
} from '../src/renderer/src/workspace-ai/panel-client'
import { workspaceConversationKey } from '../src/shared/workspace-ai-api'
import {
  WORKSPACE_PANEL_CHANNELS as C,
  type WorkspacePanelRequest,
  type WorkspacePanelSession,
  type WorkspacePanelSource,
} from '../src/shared/workspace-panel-api'
import {
  workspaceLayout,
  type WorkspaceLayout,
  type WorkspacePanelLayout,
} from '../src/shared/workspace-layout'
import type { TabManager } from '../src/main/tab-manager'
import { deferred, makeWorkspaceHost, until } from './helpers/workspace-host-fixture'
import type { ProjectFolderServices } from '../src/main/projects/project-ipc'
import type { WorkspaceAttachmentServices } from '../src/main/workspace-ai/attachments'
import type { ProjectContextBundle } from '@genoffice/project-store'

const cleanups: Array<() => void | Promise<void>> = []
afterEach(async () => {
  for (const off of cleanups.splice(0).reverse()) await off()
})

function fixture() {
  const root = mkdtempSync(join(tmpdir(), 'workspace-panel-ipc-'))
  const local = new ProjectStore(root)
  const f = makeWorkspaceHost()
  const a = f.addEditor('A')
  const b = f.addEditor('B')
  const aRequestOriginal = a.client.request.bind(a.client)
  const aRequest = vi.spyOn(a.client, 'request')
  const bRequest = vi.spyOn(b.client, 'request')
  const editors = new Map([
    ['A', a],
    ['B', b],
  ])
  const renderer = new EventEmitter()
  const contents = Object.assign(new EventEmitter(), {
    id: 10,
    mainFrame: {},
    isDestroyed: () => false,
    send: (channel: string, value: unknown) => renderer.emit(channel, {}, structuredClone(value)),
  })
  const win = Object.assign(new EventEmitter(), {
    id: 1,
    webContents: contents,
    isDestroyed: () => false,
  })
  const owner = (generation: number) => ({
    windowId: 1,
    panelContentsId: 10,
    panelGeneration: generation,
  })
  const layoutListeners = new Set<(layout: WorkspaceLayout) => void>()
  let panel: WorkspacePanelLayout | null = null
  let viewport = { width: 1280, height: 720, uiScale: 1 }
  const tabs = {
    workspaceAiForPanel: (
      generation: number,
      events: Parameters<TabManager['workspaceAiForPanel']>[1],
    ) =>
      new WorkspaceExecutionController({
        owner: owner(generation),
        ...events,
        resolve: (id) => {
          const editor = editors.get(id)
          return editor?.client.available ? { kind: 'docs', client: editor.client } : null
        },
      }),
    onWorkspaceLayoutChanged: (listener: (layout: WorkspaceLayout) => void) => {
      layoutListeners.add(listener)
      return () => {
        layoutListeners.delete(listener)
      }
    },
    setWorkspacePanel: vi.fn((next: WorkspacePanelLayout | null) => {
      panel = next
      const layout = workspaceLayout(viewport, panel)
      for (const listener of layoutListeners) listener(layout)
      return layout
    }),
    workspaceDocuments: () => [
      ...[...editors].map(([id, editor]) => ({
        id,
        kind: 'docs',
        title: id,
        active: id === 'A',
        available: editor.client.available,
      })),
      { id: 'recovery', kind: 'docs', title: 'Recovery', available: false },
      { id: 'unsupported', kind: 'pdf', title: 'PDF', available: false },
    ],
    splitState: () => null,
    tabIdForWebContents: vi.fn((_id: number): string | undefined => undefined),
    webContentsForTab: vi.fn((_id: string): unknown => undefined),
  }
  const projects = { listProjects: vi.fn(async (): Promise<ProjectListItem[]> => []) }
  const context: ProjectContextBundle = {
    projectId: 'project-one',
    name: 'Project one',
    trust: 'use',
    instructions: { text: 'PROJECT INSTRUCTIONS', sources: ['agents'] },
    memory: null,
    files: [],
    filesOmitted: 0,
    canRemember: true,
    notices: [],
  }
  const projectRuntime = {
    guide: {
      ...projects,
      readContext: vi.fn(async () => ({ status: 'ok', context: structuredClone(context) })),
      aiVisibility: vi.fn(async () => ({ root: '/project-one', ruleKey: 'visible' })),
      editMemory: vi.fn(async () => ({ status: 'ok', change: { op: 'add', text: 'memory' } })),
    },
    tools: {
      openScope: vi.fn(async () => 'private-scope'),
      closeScope: vi.fn(),
      listFiles: vi.fn(async () => ({ files: [], total: 0, offset: 0 })),
      search: vi.fn(),
      read: vi.fn(async () => ({
        status: 'ok',
        text: 'PROJECT MATERIAL',
        totalChars: 16,
        offset: 0,
      })),
    },
    scopes: { projectOf: vi.fn(() => context.projectId) },
  }
  const chats = {
    listChats: vi.fn(async (): Promise<ProjectChatListItem[]> => []),
    openChat: vi.fn(),
    append: vi.fn(),
    newChat: vi.fn(),
  }
  const history = new WorkspaceConversationHistory(local, chats, 'window-one')
  const liveSources: WorkspacePanelSource[] = []
  const sources = vi.fn(() =>
    workspacePanelSources(local, projects, chats, [
      ...liveSources,
      ...history.memoryOrigins().map((origin) => ({ origin, title: '' })),
    ]),
  )
  const sender = vi.fn(() => ({ editor: 'docs' as const, docPath: join(root, 'actual.docx') }))
  const handlers = new Map<string, Parameters<IpcMain['handle']>[1]>()
  const ipc = {
    handle: (channel: string, handler: Parameters<IpcMain['handle']>[1]) => {
      handlers.set(channel, handler)
    },
    removeHandler: (channel: string) => {
      handlers.delete(channel)
    },
  }
  const event = {
    sender: contents,
    senderFrame: contents.mainFrame,
  } as unknown as IpcMainInvokeEvent
  const attachments: WorkspaceAttachmentServices = {
    pick: vi.fn(async () => null),
    add: vi.fn((paths) => ({
      accepted: paths.map((path) => ({
        path,
        name: path.split('/').at(-1)!,
        ext: path.split('.').at(-1)!,
        sizeBytes: statSync(path).size,
      })),
      rejected: [],
    })),
    read: vi.fn(async (path, offset, max) => {
      const text = readFileSync(path, 'utf8')
      return { ok: true, offset, totalChars: text.length, text: text.slice(offset, offset + max) }
    }),
    image: vi.fn(() => ({ ok: true, base64: 'cGl4ZWw=', mime: 'image/png' })),
  }
  const drive = {
    status: vi.fn(() => ({ connected: true, enabled: true, allowed: true })),
    search: vi.fn(async () => ({
      ok: true,
      hits: [{ id: 'report', name: 'Report', type: 'document', modified: '' }],
    })),
    read: vi.fn(async () => ({
      ok: true,
      page: { name: 'Report', text: 'PRIVATE DRIVE BODY', totalChars: 18, offset: 0 },
    })),
    pick: vi.fn(async () => ({ ok: false as const, reason: 'cancelled' as const })),
    isDrivePath: vi.fn(() => false),
  }
  const openTemplate = vi.fn(async () => ({ ok: true, tabId: 'copy' }))
  const stream = vi.fn(async (_request: unknown, _signal: AbortSignal, _send: unknown) => {})
  const service = registerWorkspacePanelIpc(ipc, (event) =>
    event.sender === (contents as unknown)
      ? {
          window: win as unknown as BrowserWindow,
          tabs: tabs as unknown as TabManager,
          history,
          openTemplate,
          attachments,
          drive,
          stream,
          sources,
          sender,
          projects: projectRuntime as unknown as Pick<
            ProjectFolderServices,
            'guide' | 'tools' | 'scopes'
          >,
        }
      : null,
  )
  let delayed: { kind: string; wait: ReturnType<typeof deferred> } | null = null
  const invoke = async (channel: string, ...args: unknown[]) => {
    const result: unknown = await handlers.get(channel)!(event, ...structuredClone(args))
    if (channel === C.request && (args[1] as WorkspacePanelRequest).kind === delayed?.kind)
      await delayed.wait.promise
    return structuredClone(result)
  }
  const api = workspacePanelBridge({
    invoke,
    on: (channel, listener) => renderer.on(channel, listener),
    removeListener: (channel, listener) => renderer.removeListener(channel, listener),
  })
  const memoryService = registerWorkspaceMemoryIpc(
    ipc,
    (from) =>
      from.sender === event.sender ? { window: win as unknown as BrowserWindow, history } : null,
    service.changed,
    100,
  )
  const memoryApi = workspaceMemoryBridge({
    invoke,
    on: (channel, listener) => renderer.on(channel, listener),
    removeListener: (channel, listener) => renderer.removeListener(channel, listener),
  })
  const raw = async (request: unknown, info: WorkspacePanelSession, from = event) =>
    handlers.get(C.request)!(from, info.id, request)
  cleanups.push(() => {
    service.dispose()
    memoryService.dispose()
    f.dispose()
    rmSync(root, { recursive: true, force: true })
  })
  return {
    ...f,
    root,
    local,
    a,
    b,
    aRequest,
    aRequestOriginal,
    bRequest,
    contents,
    win,
    tabs,
    api,
    service,
    openTemplate,
    raw,
    event,
    handlers,
    history,
    liveSources,
    memoryApi,
    sources,
    projects,
    projectRuntime,
    attachments,
    drive,
    stream,
    context,
    chats,
    sender,
    layoutListeners,
    delay(kind: string) {
      const wait = deferred()
      delayed = { kind, wait }
      return wait
    },
    resize() {
      viewport = { width: 1040, height: 768, uiScale: 1.5 }
      const layout = workspaceLayout(viewport, panel)
      for (const listener of layoutListeners) listener(layout)
      return layout
    },
  }
}

async function open(f: ReturnType<typeof fixture>, client: WorkspacePanelClient) {
  const source = await client.create()
  await client.load(source.origin)
  const target = await client.capture('A')
  const run = await client.enqueue({
    conversationId: workspaceConversationKey(source.origin),
    runId: 'run',
    targetId: target.id,
    referenceIds: [],
    project: null,
  })
  return { source, target, run }
}

describe('private workspace shell service and cloned preload bridge', () => {
  it('transfers main-owned template requests only to an acknowledged live source view', async () => {
    const f = fixture()
    const client = await WorkspacePanelClient.connect(f.api)
    const from = {
      sender: { id: 21, mainFrame: {}, isDestroyed: () => false },
      senderFrame: {},
    } as unknown as IpcMainInvokeEvent
    Object.assign(from, { senderFrame: from.sender.mainFrame })
    const handoff = {
      id: 'template-one',
      mode: 'skill' as const,
      target: 'docs' as const,
      text: 'Fill the copy',
      chips: [{ id: 'app:one', source: 'app', title: 'One' }],
      createdAt: 1,
    }
    f.tabs.tabIdForWebContents.mockReturnValue('A')
    f.tabs.webContentsForTab.mockReturnValue(from.sender)
    expect(await f.service.offerTemplate(from, handoff)).toBe(false)
    await client.entriesReady(true)
    expect(
      await f.service.offerTemplate({ ...from, senderFrame: {} } as IpcMainInvokeEvent, handoff),
    ).toBe(false)
    const entries: Extract<
      import('../src/shared/workspace-panel-api').WorkspacePanelEvent,
      { kind: 'entry' }
    >[] = []
    client.subscribe((event) => {
      if (event.kind === 'entry') entries.push(event)
    })
    const offered = f.service.offerTemplate(from, handoff)
    await until(() => entries.length === 1)
    expect(entries[0]!.entry).toEqual({
      kind: 'template-handoff',
      id: handoff.id,
      text: handoff.text,
      chips: handoff.chips,
    })
    expect(await client.acknowledgeEntry(entries[0]!.entryId!, true)).toBe(true)
    expect(await offered).toBe(true)
    expect(await client.acknowledgeEntry(entries[0]!.entryId!, true)).toBe(false)
    const changed = f.service.offerTemplate(from, { ...handoff, id: 'changed' })
    await until(() => entries.length === 2)
    f.a.manual()
    expect(await client.acknowledgeEntry(entries[1]!.entryId!, true)).toBe(false)
    expect(await changed).toBe(false)
    const closed = f.service.offerTemplate(from, { ...handoff, id: 'closed' })
    await until(() => entries.length === 3)
    await client.entriesReady(false)
    expect(await closed).toBe(false)
    expect(f.stream).not.toHaveBeenCalled()
    await client.dispose()
  })

  it('expires an unacknowledged batch and rejects a late acceptance without starting a model', async () => {
    const f = fixture()
    const client = await WorkspacePanelClient.connect(f.api)
    const from = { sender: { id: 21, mainFrame: {}, isDestroyed: () => false }, senderFrame: {} }
    from.senderFrame = from.sender.mainFrame
    f.tabs.tabIdForWebContents.mockReturnValue('A')
    f.tabs.webContentsForTab.mockReturnValue(from.sender)
    await client.entriesReady(true)
    let entryId = ''
    client.subscribe((event) => {
      if (event.kind === 'entry') entryId = event.entryId!
    })
    vi.useFakeTimers()
    try {
      const result = f.handlers.get(ENTRY)!(from, {
        kind: 'edit-batch',
        text: 'BATCH',
        displayText: 'Two edits',
        documentId: 'A',
        revision: '1',
      })
      await vi.advanceTimersByTimeAsync(0)
      expect(entryId).not.toBe('')
      let answered = false
      void result.then(() => {
        answered = true
      })
      await vi.advanceTimersByTimeAsync(29_999)
      expect(answered).toBe(false)
      await vi.advanceTimersByTimeAsync(1)
      expect(await result).toBe('rejected')
      expect(await client.acknowledgeEntry(entryId, true)).toBe(false)
      expect(f.stream).not.toHaveBeenCalled()
    } finally {
      vi.useRealTimers()
      await client.dispose()
    }
  })
  it('acknowledges a batch only after installation in this live host and rejects stale or duplicate acknowledgements', async () => {
    const f = fixture()
    const client = await WorkspacePanelClient.connect(f.api)
    const from = { sender: { id: 21, mainFrame: {}, isDestroyed: () => false }, senderFrame: {} }
    from.senderFrame = from.sender.mainFrame
    f.tabs.tabIdForWebContents.mockReturnValue('A')
    f.tabs.webContentsForTab.mockReturnValue(from.sender)
    await client.entriesReady(true)
    const entries: Extract<
      import('../src/shared/workspace-panel-api').WorkspacePanelEvent,
      { kind: 'entry' }
    >[] = []
    client.subscribe((event) => {
      if (event.kind === 'entry') entries.push(event)
    })
    const entry = {
      kind: 'edit-batch',
      text: 'BATCH',
      displayText: 'Two edits',
      documentId: 'A',
      revision: '1',
    }
    const result = f.handlers.get(ENTRY)!(from, entry)
    await until(() => entries.length === 1)
    let answered = false
    void result.then(() => {
      answered = true
    })
    expect(answered).toBe(false)
    const id = entries[0]!.entryId!
    expect(entries[0]!.capture.binding.selection).toBeNull()
    expect(await client.acknowledgeEntry('unknown', true)).toBe(false)
    expect(await client.acknowledgeEntry(id, true)).toBe(true)
    expect(await result).toBe(true)
    expect(await client.acknowledgeEntry(id, true)).toBe(false)
    expect(f.stream).not.toHaveBeenCalled()

    const later = f.handlers.get(ENTRY)!(from, entry)
    await until(() => entries.length === 2)
    f.tabs.webContentsForTab.mockReturnValue({ ...from.sender })
    expect(await client.acknowledgeEntry(entries[1]!.entryId!, true)).toBe(false)
    expect(await later).toBe('rejected')
    f.tabs.webContentsForTab.mockReturnValue(from.sender)
    const changed = f.handlers.get(ENTRY)!(from, entry)
    await until(() => entries.length === 3)
    f.a.manual()
    expect(await client.acknowledgeEntry(entries[2]!.entryId!, true)).toBe(false)
    expect(await changed).toBe('rejected')
    expect(await f.handlers.get(ENTRY)!(from, entry)).toBe('rejected')
    const latest = { ...entry, revision: '2' }
    const closing = f.handlers.get(ENTRY)!(from, latest)
    await until(() => entries.length === 4)
    await client.entriesReady(false)
    expect(await closing).toBe('rejected')
    expect(await f.handlers.get(ENTRY)!(from, { ...entry, revision: '0' })).toBe(false)
    await client.dispose()
  })
  it('accepts editor navigation only for a ready host and its actual live Word view/frame', async () => {
    const f = fixture()
    const client = await WorkspacePanelClient.connect(f.api)
    const from = { sender: { id: 21, mainFrame: {}, isDestroyed: () => false }, senderFrame: {} }
    from.senderFrame = from.sender.mainFrame
    f.tabs.tabIdForWebContents.mockImplementation((id) => (id === 21 ? 'A' : undefined))
    f.tabs.webContentsForTab.mockReturnValue(from.sender)
    const route = (entry: unknown, event: unknown = from) => f.handlers.get(ENTRY)!(event, entry)
    const events: import('../src/shared/workspace-panel-api').WorkspacePanelEvent[] = []
    client.subscribe((event) => events.push(event))
    expect(await route({ kind: 'open' })).toBe(false)
    await client.entriesReady(true)
    expect(await route({ kind: 'preset', text: 'TEST_ENTRY', autoRun: true })).toBe(true)
    expect(events).toEqual([
      expect.objectContaining({
        kind: 'entry',
        entry: { kind: 'preset', text: 'TEST_ENTRY', autoRun: true },
        capture: expect.objectContaining({
          binding: expect.objectContaining({
            document: expect.objectContaining({ tabId: 'A', kind: 'docs' }),
          }),
        }),
      }),
    ])
    expect(f.aRequest.mock.calls.some(([op]) => op.kind === 'prepare' || op.kind === 'read')).toBe(
      false,
    )
    expect(f.stream).not.toHaveBeenCalled()
    expect(await route({ kind: 'open', tabId: 'B' })).toBe(false)
    expect(await route({ kind: 'open' }, { ...from, senderFrame: {} })).toBe(false)
    f.tabs.webContentsForTab.mockReturnValue({ ...from.sender })
    expect(await route({ kind: 'open' })).toBe(false)
    f.tabs.webContentsForTab.mockReturnValue(from.sender)
    await client.entriesReady(false)
    expect(await route({ kind: 'toggle' })).toBe(false)
    await client.dispose()
    expect(await route({ kind: 'open' })).toBe(false)
    expect(events.filter((event) => event.kind === 'entry')).toHaveLength(1)
  })

  it('does not deliver an entry when the host closes during the source capture', async () => {
    const f = fixture()
    const client = await WorkspacePanelClient.connect(f.api)
    const from = { sender: { id: 21, mainFrame: {}, isDestroyed: () => false }, senderFrame: {} }
    from.senderFrame = from.sender.mainFrame
    f.tabs.tabIdForWebContents.mockReturnValue('A')
    f.tabs.webContentsForTab.mockReturnValue(from.sender)
    await client.entriesReady(true)
    const pending = deferred()
    const request = f.aRequest.getMockImplementation()!
    f.aRequest.mockImplementation(async (op, signal) => {
      if (op.kind === 'capture') await pending.promise
      return request(op, signal)
    })
    const result = f.handlers.get(ENTRY)!(from, { kind: 'open' })
    await client.dispose()
    pending.resolve()
    expect(await result).toBe(false)
  })

  it('restricts model turns to a live run, relays structured chunks, and awaits cancellation', async () => {
    const f = fixture()
    const client = await WorkspacePanelClient.connect(f.api)
    const { run: bound } = await open(f, client)
    const request = {
      requestId: 'turn-one',
      sessionId: 'model-session',
      system: '',
      messages: [],
      tools: [],
    }
    await expect(client.stream(bound, request)).rejects.toThrow()
    const run = (await client.startNext())!
    const events: unknown[] = []
    const off = client.subscribe((event) => {
      if (event.kind === 'stream') events.push(event.chunk)
    })
    for (const bad of [
      { ...run.binding, conversationId: 'wrong' },
      { ...run.binding, owner: { ...run.binding.owner, panelGeneration: 9 } },
    ])
      await expect(client.stream(bad, request)).rejects.toThrow()
    expect(f.stream).not.toHaveBeenCalled()
    const stopped = deferred()
    let signal: AbortSignal | undefined
    f.stream.mockImplementation(async (_request, abort, send) => {
      signal = abort
      ;(send as (chunk: unknown) => void)({ requestId: request.requestId, type: 'input-truncated' })
      await stopped.promise
      ;(send as (chunk: unknown) => void)({ requestId: request.requestId, type: 'done' })
    })
    const sending = client.stream(run.binding, request)
    await until(() => f.stream.mock.calls.length === 1)
    const cancelled = vi.fn()
    const cancelling = client.cancel(run.binding).then(cancelled)
    await until(() => signal?.aborted === true)
    expect(cancelled).not.toHaveBeenCalled()
    stopped.resolve()
    await Promise.all([sending, cancelling])
    expect(cancelled).toHaveBeenCalledOnce()
    expect(events).toEqual([
      { requestId: request.requestId, type: 'input-truncated' },
      { requestId: request.requestId, type: 'done' },
    ])
    await expect(client.stream(run.binding, { ...request, requestId: 'late' })).rejects.toThrow()
    off()
    await client.dispose()
  })

  it('rechecks Drive permission at the model entry and rejects malformed turn requests', async () => {
    const f = fixture()
    const client = await WorkspacePanelClient.connect(f.api)
    await open(f, client)
    const run = (await client.startNext())!
    const request = {
      requestId: 'turn-one',
      sessionId: 'session-one',
      system: '',
      messages: [],
      tools: [],
    }
    f.drive.isDrivePath.mockReturnValue(true)
    f.drive.status.mockReturnValue({ connected: true, enabled: false, allowed: false })
    await expect(client.stream(run.binding, request)).rejects.toThrow()
    for (const changed of [
      { messages: [{ role: 'unexpected' }] },
      { tools: [{ name: 'invalid tool' }] },
      { requestId: '' },
    ]) {
      expect(
        await f.raw(
          { kind: 'stream', run: run.binding, request: { ...request, ...changed } },
          client.session,
        ),
      ).toMatchObject({ ok: false })
    }
    expect(f.stream).not.toHaveBeenCalled()
    await client.cancel(run.binding)
    await client.dispose()
  })

  it('binds picked files to the queued conversation and rejects malformed, ungranted, retired and foreign reads', async () => {
    const f = fixture()
    const path = join(f.root, 'material.md')
    writeFileSync(path, 'PRIVATE ATTACHMENT TEXT')
    const client = await WorkspacePanelClient.connect(f.api)
    await expect(client.addAttachments('unopened', [path])).rejects.toThrow()
    const source = await client.create()
    await client.load(source.origin)
    const id = workspaceConversationKey(source.origin)
    const files = (await client.addAttachments(id, [path]))!.accepted
    const target = await client.capture('A')
    const request = {
      conversationId: id,
      runId: 'attachment-run',
      targetId: target.id,
      referenceIds: [],
      project: null,
    }
    await expect(client.enqueue(request, [{ ...files[0]!, path: '/ungranted' }])).rejects.toThrow()
    const bound = await client.enqueue({ ...request, runId: 'granted-run' }, files)
    await expect(client.attachmentRead(bound, 0, 0)).rejects.toThrow()
    const run = (await client.startNext())!
    expect(await client.attachmentList(run.binding)).toEqual(files)
    expect((await client.attachmentRead(run.binding, 0, 0)).text).toBe('PRIVATE ATTACHMENT TEXT')
    for (const bad of [
      { ...run.binding, conversationId: 'other' },
      { ...run.binding, owner: { ...run.binding.owner, panelGeneration: 99 } },
    ])
      await expect(client.attachmentRead(bad, 0, 0)).rejects.toThrow()
    await expect(client.attachmentRead(run.binding, -1, 0)).rejects.toThrow()
    await expect(client.attachmentRead(run.binding, 0, 0.5)).rejects.toThrow()
    await expect(client.attachmentImage(run.binding, 0)).rejects.toThrow()
    await client.cancel(run.binding)
    await expect(client.attachmentRead(run.binding, 0, 0)).rejects.toThrow()
    expect(f.attachments.read).toHaveBeenCalledTimes(1)
    await client.dispose()
  })

  it('prepares only the sent attachments and excludes their tool text from saved history', async () => {
    const f = fixture()
    const path = join(f.root, 'material.md')
    writeFileSync(path, 'ATTACHMENT BODY')
    const requests: AgentStreamRequest[] = []
    const callbacks: AgentStreamCallbacks[] = []
    const connected = await connectWorkspacePanel(f.api, {
      gate: () => ({ kind: 'run' }),
      admit: () => true,
      prepare: async () => ({ project: null, skills: [], dispose() {} }),
      errorText: String,
      createTransport: () => ({
        stream(request, cb) {
          requests.push(request)
          callbacks.push(cb)
          return { cancel: () => cb.onDone() }
        },
      }),
    })
    cleanups.push(() => connected.dispose())
    await connected.controller.create()
    const first = connected.host.conversations.selectedId!
    await connected.controller.target(first, 'A')
    const files = (await connected.client.addAttachments(first, [path]))!.accepted
    const composer = connected.host.conversations.state(first).composer
    connected.host.conversations.edit(first, {
      composer: { ...composer, text: 'Read material', attachments: files },
    })
    await connected.controller.send(first)
    await until(() => requests.length === 1)
    expect(JSON.stringify(requests[0])).toContain('material.md')
    callbacks[0]!.onToolCall({
      id: 'attachment-call',
      name: 'read_attachment',
      input: { index: 0 },
    })
    callbacks[0]!.onDone()
    await until(() => requests.length === 2)
    expect(JSON.stringify(requests[1])).toContain('ATTACHMENT BODY')
    await connected.controller.create()
    const second = connected.host.conversations.selectedId!
    expect(connected.host.conversations.state(second).composer.attachments).toEqual([])
    callbacks[1]!.onDelta('Read the attachment')
    callbacks[1]!.onDone()
    await until(() => !connected.host.busy())
    const history = await connected.client.load(connected.host.conversations.state(first).origin)
    expect(history.messages[0]!.attachments).toEqual(files)
    expect(JSON.stringify(history)).not.toContain('ATTACHMENT BODY')
    expect(connected.host.conversations.selectedId).toBe(second)
  })

  it('creates project conversations without reading context, then binds guarded project tools to the explicit send', async () => {
    const f = fixture()
    f.projects.listProjects.mockResolvedValue([
      { id: 'project-one', name: 'Project one', missing: false } as ProjectListItem,
    ])
    f.chats.newChat.mockResolvedValue({ chatId: 'new-project-chat' })
    f.chats.openChat.mockResolvedValue({ status: 'missing' })
    f.chats.append.mockResolvedValue({ status: 'saved' })
    const requests: AgentStreamRequest[] = []
    let cb!: AgentStreamCallbacks
    const connected = await connectWorkspacePanel(f.api, {
      gate: () => ({ kind: 'run' }),
      admit: () => true,
      prepare: async () => ({ project: null, skills: [], dispose() {} }),
      errorText: String,
      createTransport: () => ({
        stream(request, callbacks) {
          requests.push(request)
          cb = callbacks
          return { cancel: () => callbacks.onDone() }
        },
      }),
    })
    cleanups.push(() => connected.dispose())
    await connected.controller.create()
    const local = connected.host.conversations.selectedId!
    await connected.controller.create('project-one')
    const id = connected.host.conversations.selectedId!
    expect(id).not.toBe(local)
    expect(connected.host.conversations.state(id).origin).toEqual({
      kind: 'project',
      projectId: 'project-one',
      chatId: 'new-project-chat',
    })
    expect(f.projectRuntime.guide.readContext).not.toHaveBeenCalled()
    expect(requests).toEqual([])
    connected.host.conversations.edit(id, {
      composer: { ...connected.host.conversations.state(id).composer, text: 'Use this project' },
    })
    await connected.controller.send(id)
    await until(() => requests.length === 1)
    expect(requests[0]!.system).toContain('PROJECT INSTRUCTIONS')
    expect(JSON.stringify(requests[0]!.messages)).not.toContain('PROJECT INSTRUCTIONS')
    expect(requests[0]!.tools.map((t) => t.name)).toContain('project_read')
    const binding = connected.host.conversations.state(id).runs[0]!.input.project!
    expect(JSON.stringify(binding)).not.toContain('private-scope')
    cb.onToolCall({ id: 'project-read', name: 'project_read', input: { path: 'material.md' } })
    cb.onDone()
    await until(() => f.projectRuntime.tools.read.mock.calls.length === 1)
    expect(f.projectRuntime.tools.read).toHaveBeenCalledWith(10, {
      relPath: 'material.md',
      scope: 'private-scope',
    })
    connected.controller.select(local)
    const run = connected.host.conversations.state(id).runs[0]!.run
    await connected.host.cancel(run)
    expect(f.projectRuntime.tools.closeScope).toHaveBeenCalledWith(10, { scope: 'private-scope' })
    expect(connected.host.conversations.selectedId).toBe(local)
    expect(connected.host.conversations.state(local).messages).toEqual([])
    await expect(
      connected.client.projectTool(
        {
          owner: connected.client.session.owner,
          conversationId: id,
          runId: run.runId,
          bindingId: 1,
        },
        'read',
        { path: 'material.md' },
      ),
    ).rejects.toThrow('invalid-request')
  })

  it('rejects arbitrary project capabilities and raw scopes before service calls', async () => {
    const f = fixture()
    const client = await WorkspacePanelClient.connect(f.api)
    const { run } = await open(f, client)
    await expect(
      client.enqueue({
        conversationId: run.conversationId,
        runId: 'forged-project',
        targetId: (await client.capture('A')).id,
        referenceIds: [],
        project: { id: 'project-one', contextVersion: 'forged' },
      }),
    ).rejects.toThrow('invalid-request')
    expect(
      await f.raw(
        {
          kind: 'project-tool',
          run,
          method: 'read',
          args: { path: 'material.md', scope: 'private-scope' },
        },
        client.session,
      ),
    ).toEqual({ ok: false, reason: 'invalid-request' })
    expect(
      await f.raw(
        {
          kind: 'project-tool',
          run,
          method: 'remember',
          args: { change: { op: 'replace', old: 'one' } },
        },
        client.session,
      ),
    ).toEqual({ ok: false, reason: 'invalid-request' })
    expect(f.projectRuntime.guide.readContext).not.toHaveBeenCalled()
    expect(f.projectRuntime.tools.read).not.toHaveBeenCalled()
  })

  it('refuses changed project context in the FIFO queue and restores the original draft', async () => {
    const f = fixture()
    f.projects.listProjects.mockResolvedValue([
      { id: 'project-one', name: 'Project one', missing: false } as ProjectListItem,
    ])
    f.chats.newChat.mockResolvedValue({ chatId: 'queued-project-chat' })
    f.chats.openChat.mockResolvedValue({ status: 'missing' })
    const callbacks: AgentStreamCallbacks[] = []
    const connected = await connectWorkspacePanel(f.api, {
      gate: () => ({ kind: 'run' }),
      admit: () => true,
      prepare: async () => ({ project: null, skills: [], dispose() {} }),
      errorText: String,
      createTransport: () => ({
        stream(_request, cb) {
          callbacks.push(cb)
          return { cancel: () => cb.onDone() }
        },
      }),
    })
    cleanups.push(() => connected.dispose())
    await connected.controller.create()
    const first = connected.host.conversations.selectedId!
    connected.host.conversations.edit(first, {
      composer: { ...connected.host.conversations.state(first).composer, text: 'First request' },
    })
    await connected.controller.send(first)
    await until(() => callbacks.length === 1)
    await connected.controller.create('project-one')
    const second = connected.host.conversations.selectedId!
    connected.host.conversations.edit(second, {
      composer: {
        ...connected.host.conversations.state(second).composer,
        text: 'Review this project',
      },
    })
    await connected.controller.send(second)
    expect(f.projectRuntime.guide.readContext).toHaveBeenCalledOnce()
    expect(f.projectRuntime.tools.openScope).not.toHaveBeenCalled()
    f.context.instructions!.text = 'CHANGED IN QUEUE'
    callbacks[0]!.onDelta('First reply')
    callbacks[0]!.onDone()
    await until(() => connected.host.conversations.state(second).runs.length === 0)
    expect(callbacks).toHaveLength(1)
    expect(connected.host.conversations.state(second).composer.text).toBe('Review this project')
    expect(connected.host.conversations.state(second).error).not.toBeNull()
    expect(f.projectRuntime.tools.openScope).not.toHaveBeenCalled()
  })
  it('routes home memory through the original host, releases its lease and never creates a chat file', async () => {
    const f = fixture()
    const messages = [
      { role: 'assistant' as const, text: 'original home', seq: 0, ts: new Date().toISOString() },
    ]
    let locked = false
    let busy = true
    const off = await f.memoryApi.register(
      homeMemoryHost({
        busy: () => busy,
        messages: () => messages,
        append: (message) => {
          messages.push(message as (typeof messages)[number])
        },
        locked: (value) => {
          locked = value
        },
      }),
    )
    const client = await WorkspacePanelClient.connect(f.api)
    const source = (await client.catalog()).sources.find((s) => s.origin.kind === 'window-memory')!
    expect(source).toBeDefined()
    await expect(client.load(source.origin)).rejects.toThrow('service-error')
    expect(locked).toBe(false)
    busy = false
    expect((await client.load(source.origin)).messages[0]?.text).toBe('original home')
    expect(locked).toBe(true)
    await client.append(source.origin, { role: 'assistant', text: 'from workspace' })
    expect(messages.at(-1)?.text).toBe('from workspace')
    expect(readdirSync(f.root)).toEqual([])
    await client.dispose()
    expect(locked).toBe(false)
    off()
    const next = await WorkspacePanelClient.connect(f.api)
    expect((await next.catalog()).sources).toEqual([])
    await expect(next.load(source.origin)).rejects.toThrow('invalid-request')
  })

  it('rejects cloned contents and subframes for the original home host and its replies', async () => {
    const f = fixture()
    const call = (channel: string, event = f.event, ...args: unknown[]) =>
      f.handlers.get(channel)!(event, ...args)
    for (const wrong of [
      { sender: f.contents, senderFrame: {} },
      { sender: { ...f.contents }, senderFrame: f.contents.mainFrame },
    ])
      expect(await call(M.register, wrong as IpcMainInvokeEvent)).toBeNull()
    const id = await call(M.register)
    let query: { id: string } | undefined
    f.contents.send = (channel, value) => {
      if (channel === M.request) query = value as { id: string }
    }
    const origin = f.history.memoryOrigins()[0]!
    const pending = f.history.load(origin)
    await until(() => !!query)
    const answer = {
      ok: true,
      messages: [{ role: 'assistant', text: 'accepted', seq: 0, ts: new Date().toISOString() }],
    }
    await call(
      M.reply,
      {
        sender: { ...f.contents },
        senderFrame: f.contents.mainFrame,
      } as unknown as IpcMainInvokeEvent,
      id,
      query!.id,
      answer,
    )
    await call(
      M.reply,
      { ...f.event, senderFrame: {} } as IpcMainInvokeEvent,
      id,
      query!.id,
      answer,
    )
    await call(M.reply, f.event, id, query!.id, answer)
    expect((await pending).messages[0]?.text).toBe('accepted')
    const abandoned = f.history.load(origin)
    const rejected = expect(abandoned).rejects.toThrow('unavailable')
    f.contents.emit('did-start-navigation', {}, 'reloaded', false, true)
    await rejected
    expect(f.history.memoryOrigins()).toEqual([])
    await call(M.reply, f.event, id, query!.id, answer)
    expect(readdirSync(f.root)).toEqual([])
  })

  it('records the started target even when the next queued request selects a different document', async () => {
    const f = fixture()
    const client = await WorkspacePanelClient.connect(f.api)
    const { source, run } = await open(f, client)
    const first = (await client.startNext())!
    await client.complete(first.binding)
    const target = await client.capture('B')
    await client.enqueue({
      conversationId: run.conversationId,
      runId: 'next-target',
      targetId: target.id,
      referenceIds: [],
      project: null,
    })
    await client.append(source.origin, { role: 'assistant', text: 'first answer' })
    expect(f.sender).toHaveBeenLastCalledWith('A')
    const second = (await client.startNext())!
    await client.complete(second.binding)
    await client.append(source.origin, { role: 'assistant', text: 'second answer' })
    expect(f.sender).toHaveBeenLastCalledWith('B')
    await client.dispose()
  })

  it('follows unsaved first-save references and accepts no orphan writes after a rebind', async () => {
    const f = fixture()
    f.local.ensureDefaultProject()
    const origin = { kind: 'legacy-document' as const, projectId: 'default', chatId: 'unsaved-one' }
    f.liveSources.push({ origin, title: 'Untitled' })
    f.local.appendChatMessage('default', 'unsaved-one', { role: 'user', text: 'pending original' })
    const client = await WorkspacePanelClient.connect(f.api)
    expect((await client.load(origin)).messages[0]?.text).toBe('pending original')
    const ref = f.local.rebindChatToFile('default', 'unsaved-one', join(f.root, 'saved.docx'))
    const saved = { kind: 'legacy-document' as const, ...ref }
    f.liveSources[0] = { origin: saved, title: 'saved.docx' }
    f.service.rebound(origin, saved)
    await client.append(origin, { role: 'assistant', text: 'after save' })
    expect(f.local.loadChat(ref.projectId, ref.chatId).map((m) => m.text)).toEqual([
      'pending original',
      'after save',
    ])
    expect(f.local.loadChat('default', 'unsaved-one')).toEqual([])
    expect((await client.load(origin)).origin).toEqual(saved)
    expect(f.local.listWorkspaceChats()).toEqual([])
    await client.dispose()
  })

  it('lists actual sources and metadata without reading a document; references return handles only', async () => {
    const f = fixture()
    const client = await WorkspacePanelClient.connect(f.api)
    expect((await client.catalog()).documents).toHaveLength(4)
    expect(f.aRequest).not.toHaveBeenCalled()
    expect(f.bRequest).not.toHaveBeenCalled()
    expect(readdirSync(f.root)).toEqual([])
    const { run } = await open(f, client)
    const capture = await client.capture('B')
    const reference = await client.reference(capture.id, 'document')
    expect(reference).toHaveProperty('binding.document.tabId', 'B')
    expect(reference).not.toHaveProperty('snapshot')
    expect(JSON.stringify(reference)).not.toContain('B UNSAVED')
    expect(f.bRequest.mock.calls.map(([op]) => op.kind)).toEqual(['describe', 'read'])
    await client.cancel(run)
    await client.dispose()
  })

  it('rejects other contents, subframes, stale sessions, malformed requests and caller-supplied project bindings', async () => {
    const f = fixture()
    const client = await WorkspacePanelClient.connect(f.api)
    const { run, target } = await open(f, client)
    for (const event of [
      { sender: f.contents, senderFrame: {} },
      { sender: { ...f.contents }, senderFrame: f.contents.mainFrame },
    ])
      expect(
        await f.raw({ kind: 'capture', tabId: 'B' }, client.session, event as IpcMainInvokeEvent),
      ).toMatchObject({ ok: false, reason: 'wrong-owner' })
    const invalid: unknown[] = [
      { kind: 'drive-pick', conversationId: 'unopened' },
      { kind: 'drive-check', run: { ...run, owner: null } },
      { kind: 'drive-search', run, query: 'x', limit: 21 },
      { kind: 'drive-search', run, query: 'x', limit: 0.5 },
      { kind: 'drive-read', run, fileId: '../outside', offset: 0 },
      { kind: 'drive-read', run, fileId: 'report', offset: -1 },
      {
        kind: 'execute',
        args: [run, target.binding.document, { id: 'x', name: 'write', input: null }],
      },
      { kind: 'layout', panel: { side: 'left', width: NaN, open: true } },
      { kind: 'reference', captureId: target.id, scope: 'all-files' },
      {
        kind: 'enqueue',
        request: {
          conversationId: run.conversationId,
          runId: 'bad',
          targetId: target.id,
          referenceIds: [],
          project: { id: 'forged', contextVersion: 'forged' },
        },
      },
      { kind: 'load', origin: { kind: 'workspace', conversationId: '../outside' } },
      {
        kind: 'append',
        origin: { kind: 'workspace', conversationId: 'unknown' },
        message: { role: 'system', text: 'no' },
      },
      { kind: 'executeJavaScript', source: 'no' },
    ]
    for (const request of invalid)
      expect(await f.raw(request, client.session)).toMatchObject({
        ok: false,
        reason: 'invalid-request',
      })
    expect(f.a.body()).toBe('A UNSAVED')
    expect(f.bRequest).not.toHaveBeenCalled()
    const next = await WorkspacePanelClient.connect(f.api)
    expect(next.session.owner.panelGeneration).toBeGreaterThan(client.session.owner.panelGeneration)
    expect(await f.raw({ kind: 'cancel', args: [run] }, client.session)).toMatchObject({
      ok: false,
      reason: 'wrong-owner',
    })
    expect(await f.raw({ kind: 'cancel', args: [run] }, next.session)).toMatchObject({
      ok: false,
      reason: 'wrong-owner',
    })
    await expect(client.catalog()).rejects.toThrow('closed')
    await next.dispose()
  })

  it('keeps shared buffered history and stable legacy references without creating a second store', async () => {
    const f = fixture()
    f.local.ensureDefaultProject()
    const ref = f.local.resolveChatForFile(join(f.root, 'saved.docx'))
    f.local.appendChatMessage(ref.projectId, ref.chatId, {
      role: 'user',
      text: 'buffered in original store',
    })
    // Pending opening records are not yet catalogued; materialize the first reply through the same owner.
    f.local.appendChatMessage(ref.projectId, ref.chatId, {
      role: 'assistant',
      text: 'original answer',
    })
    const client = await WorkspacePanelClient.connect(f.api)
    const source = (await client.catalog()).sources.find(
      (s) => s.origin.kind === 'legacy-document',
    )!
    const before = readFileSync(
      join(f.root, 'projects', ref.projectId, 'chats', `${ref.chatId}.jsonl`),
      'utf8',
    )
    expect((await client.load(source.origin)).messages.map((m) => m.text)).toEqual([
      'buffered in original store',
      'original answer',
    ])
    f.local.fileRenamed(join(f.root, 'saved.docx'), join(f.root, 'renamed.docx'))
    await client.append(source.origin, { role: 'assistant', text: 'continued' })
    expect(f.local.loadChat(ref.projectId, ref.chatId).at(-1)?.text).toBe('continued')
    expect(f.local.listWorkspaceChats()).toEqual([])
    expect(
      readFileSync(join(f.root, 'projects', ref.projectId, 'chats', `${ref.chatId}.jsonl`), 'utf8'),
    ).toContain(before)
    const created = await client.create()
    await client.load(created.origin)
    await client.append(created.origin, { role: 'user', text: 'pending user' })
    expect((await f.history.load(created.origin)).messages.at(-1)?.text).toBe('pending user')
    await client.append(created.origin, { role: 'assistant', text: 'new answer' })
    const restarted = new WorkspaceConversationHistory(new ProjectStore(f.root), f.chats, 'restart')
    expect((await restarted.load(created.origin)).messages.map((m) => m.text)).toEqual([
      'pending user',
      'new answer',
    ])
    await client.dispose()
  })

  it('uses project visibility and write refusal without accepting unknown origins or saving a local substitute', async () => {
    const f = fixture()
    f.projects.listProjects.mockResolvedValue([
      { id: 'project', name: 'Project', missing: false } as ProjectListItem,
    ])
    f.chats.listChats.mockResolvedValue([
      { id: 'visible', title: 'Visible' } as ProjectChatListItem,
    ])
    f.chats.openChat.mockResolvedValue({ status: 'ok', messages: [], foreign: true })
    f.chats.append.mockResolvedValue({ status: 'unwritable' })
    const client = await WorkspacePanelClient.connect(f.api)
    const sources = (await client.catalog()).sources
    expect(sources).toHaveLength(1)
    expect((await client.load(sources[0]!.origin)).foreign).toBe(true)
    expect(
      await client.append(sources[0]!.origin, { role: 'assistant', text: 'not saved locally' }),
    ).toBe('unwritable')
    expect(readdirSync(f.root)).toEqual([])
    await expect(
      client.load({ kind: 'project', projectId: 'project', chatId: 'hidden' }),
    ).rejects.toThrow('invalid-request')
    f.chats.listChats.mockResolvedValue([])
    await expect(
      client.append(sources[0]!.origin, { role: 'user', text: 'after revocation' }),
    ).rejects.toThrow('invalid-request')
    expect(f.chats.append).toHaveBeenCalledTimes(1)
    await client.dispose()
  })

  it('restores abort signals when a stop overtakes the start reply and refuses late replies after navigation', async () => {
    const f = fixture()
    const client = await WorkspacePanelClient.connect(f.api)
    const { run } = await open(f, client)
    const wait = f.delay('startNext')
    const starting = client.startNext()
    await until(() => f.aRequest.mock.calls.some(([op]) => op.kind === 'prepare'))
    f.a.manual()
    wait.resolve()
    const started = await starting
    expect(started?.signal.aborted).toBe(true)
    expect(await client.state(run.runId)).toMatchObject({ status: 'stopped' })
    const pending = f.delay('capture')
    const capture = client.capture('A')
    await until(() => f.aRequest.mock.calls.filter(([op]) => op.kind === 'describe').length >= 2)
    f.contents.emit('did-start-navigation', {}, 'about:blank', false, true)
    pending.resolve()
    await expect(capture).rejects.toThrow('closed')
    expect(f.layoutListeners.size).toBe(0)
    expect(f.tabs.setWorkspacePanel).toHaveBeenLastCalledWith(null)
  })

  it('forwards main layout changes and cleans reservations on disconnect without touching split preferences', async () => {
    const f = fixture()
    const client = await WorkspacePanelClient.connect(f.api)
    const events: unknown[] = []
    client.subscribe((event) => events.push(event))
    const initial = await client.layout({ side: 'left', width: 360, open: true })
    expect(initial.documents.x).toBe(360)
    const resized = f.resize()
    expect(events.at(-1)).toMatchObject({ kind: 'layout', layout: resized })
    expect(resized.documents.x).toBe(540)
    await client.dispose()
    expect(f.tabs.setWorkspacePanel).toHaveBeenLastCalledWith(null)
    expect(f.layoutListeners.size).toBe(0)
    expect(f.contents.listenerCount('destroyed')).toBe(0)
  })

  it('connects panel selectors and the real run host through cloned IPC, keeping replies, edits and rollback in the original conversation', async () => {
    const f = fixture()
    const requests: AgentStreamRequest[] = []
    const callbacks: AgentStreamCallbacks[] = []
    const cancelled = vi.fn()
    const gate = aiSendGate(
      { ...defaultAiSettings(), subscriptionNoticeSeen: true },
      { 'claude-code': { provider: 'claude-code', state: 'ready', meteredAccepted: false } },
    )
    const connected = await connectWorkspacePanel(f.api, {
      gate: () => gate,
      admit: () => gate.kind === 'run',
      prepare: async () => ({ project: null, skills: [], dispose() {} }),
      createTransport: () => ({
        stream(request, cb) {
          requests.push(request)
          callbacks.push(cb)
          return {
            cancel() {
              cancelled()
              queueMicrotask(() => cb.onDone())
            },
          }
        },
      }),
      errorText: String,
    })
    cleanups.push(connected.dispose)
    await connected.controller.create()
    const p = connected.host.conversations.selectedId!
    await connected.controller.addReference(p, 'B', 'document')
    connected.host.conversations.edit(p, {
      composer: { ...connected.host.conversations.state(p).composer, text: 'Use B to edit A' },
    })
    await connected.controller.send(p)
    await until(() => requests.length === 1)
    expect(JSON.stringify(requests[0])).toContain('B UNSAVED')
    await connected.controller.create()
    const q = connected.host.conversations.selectedId!
    callbacks[0]!.onToolCall({ id: 'one', name: 'write', input: { text: 'A from B' } })
    callbacks[0]!.onDone()
    await until(() => requests.length === 2)
    callbacks[1]!.onDelta('Finished P')
    callbacks[1]!.onDone()
    await until(() => !connected.host.busy())
    expect(f.a.body()).toBe('A from B')
    expect(f.b.body()).toBe('B UNSAVED')
    expect(connected.host.conversations.state(q).messages).toEqual([])
    expect(connected.host.conversations.state(p).messages.at(-1)?.text).toBe('Finished P')
    const card = connected.host.rollbackCards(p)[0]!
    await connected.host.rollback(p, card.id)
    expect(f.a.body()).toBe('A UNSAVED')
    expect(f.sender).toHaveBeenCalledWith('A')
    expect(cancelled).not.toHaveBeenCalled()
    await connected.controller.target(p, 'A')
    connected.host.conversations.edit(p, {
      composer: { ...connected.host.conversations.state(p).composer, text: 'Second request' },
    })
    await connected.controller.send(p)
    await until(() => requests.length === 3)
    f.b.manual()
    await until(() => cancelled.mock.calls.length === 1 && !connected.host.busy())
    callbacks[2]!.onToolCall({ id: 'late', name: 'write', input: { text: 'must not apply' } })
    callbacks[2]!.onDone()
    expect(f.a.body()).toBe('A UNSAVED')
  })
})

describe('workspace Drive through the production IPC bridge', () => {
  it('binds reads to the running conversation, blocks outbound tools and omits fetched bodies from history', async () => {
    const f = fixture()
    const requests: AgentStreamRequest[] = []
    const callbacks: AgentStreamCallbacks[] = []
    const outbound = vi.fn(async () => ({ output: 'must not run', summary: 'web' }))
    const connected = await connectWorkspacePanel(f.api, {
      gate: () => ({ kind: 'run' }),
      admit: () => true,
      errorText: String,
      prepare: async () => ({
        project: null,
        skills: [
          {
            id: 'web',
            systemPrompt: '',
            tools: [{ name: 'web_search', description: 'search', inputSchema: { type: 'object' } }],
            executeTool: outbound,
          },
        ],
        dispose() {},
      }),
      createTransport: () => ({
        stream(request, cb) {
          requests.push(request)
          callbacks.push(cb)
          return { cancel: () => cb.onDone() }
        },
      }),
    })
    cleanups.push(() => connected.dispose())
    await connected.controller.create()
    const id = connected.host.conversations.selectedId!
    await connected.controller.target(id, 'A')
    expect(f.drive.read).not.toHaveBeenCalled()
    expect(requests).toEqual([])
    connected.host.conversations.edit(id, {
      composer: { ...connected.host.conversations.state(id).composer, text: 'Read report' },
    })
    await connected.controller.send(id)
    await until(() => requests.length === 1)
    callbacks[0]!.onToolCall({ id: 'read-drive', name: 'drive_read', input: { file_id: 'report' } })
    callbacks[0]!.onDone()
    await connected.controller.create()
    const other = connected.host.conversations.selectedId!
    await until(() => requests.length === 2)
    expect(JSON.stringify(requests[1])).toContain('PRIVATE DRIVE BODY')
    callbacks[1]!.onToolCall({ id: 'web', name: 'web_search', input: { query: 'private' } })
    callbacks[1]!.onDone()
    await until(() => requests.length === 3)
    expect(outbound).not.toHaveBeenCalled()
    expect(JSON.stringify(requests[2])).toContain('content read from Google Drive')
    callbacks[2]!.onDelta('Finished Drive request')
    callbacks[2]!.onDone()
    await until(() => !connected.host.busy())
    const history = await connected.client.load(connected.host.conversations.state(id).origin)
    expect(JSON.stringify(history)).not.toContain('PRIVATE DRIVE BODY')
    expect(history.messages.at(-1)?.tools?.some((t) => t.name === 'drive_read')).toBe(true)
    expect(connected.host.conversations.state(other).messages).toEqual([])
    f.drive.status.mockReturnValue({ connected: true, enabled: false, allowed: false })
    await connected.controller.target(id, 'A')
    connected.host.conversations.edit(id, {
      composer: {
        ...connected.host.conversations.state(id).composer,
        text: 'Continue with report',
      },
    })
    await connected.controller.send(id)
    await until(() => !connected.host.busy())
    expect(requests).toHaveLength(3)
    expect(connected.host.conversations.state(id).composer.text).toBe('Continue with report')
  })
})

describe('project template service ownership', () => {
  it('derives copy placement from the original conversation and validates its target before opening', async () => {
    const f = fixture()
    f.projects.listProjects.mockResolvedValue([
      { id: 'project-one', name: 'Project one', missing: false } as ProjectListItem,
    ])
    f.chats.newChat.mockResolvedValue({ chatId: 'project-copy-chat' })
    f.chats.openChat.mockResolvedValue({ status: 'missing' })
    const client = await WorkspacePanelClient.connect(f.api)
    const project = await client.create('project-one')
    const local = await client.create()
    await client.load(project.origin)
    await client.load(local.origin)
    const input = { skillId: 'app:one', text: 'Fill copy', chips: [] }
    expect(await client.openTemplate(workspaceConversationKey(project.origin), input, 'A')).toEqual(
      { ok: true, tabId: 'copy' },
    )
    expect(f.openTemplate).toHaveBeenLastCalledWith(input, 'project-one', 'A')
    await client.openTemplate(workspaceConversationKey(local.origin), input)
    expect(f.openTemplate).toHaveBeenLastCalledWith(input, null, undefined)
    await expect(
      client.openTemplate(workspaceConversationKey(project.origin), input, 'recovery'),
    ).rejects.toBeDefined()
    await expect(client.openTemplate('unopened', input)).rejects.toBeDefined()
    expect(f.openTemplate).toHaveBeenCalledTimes(2)
    expect(f.projectRuntime.guide.readContext).not.toHaveBeenCalled()
    await client.dispose()
  })

  it('requires an acknowledged project copy capability before completion can save it', async () => {
    const f = fixture()
    f.projects.listProjects.mockResolvedValue([
      { id: 'project-one', name: 'Project one', missing: false } as ProjectListItem,
    ])
    f.chats.newChat.mockResolvedValue({ chatId: 'project-copy-chat' })
    f.chats.openChat.mockResolvedValue({ status: 'missing' })
    const client = await WorkspacePanelClient.connect(f.api)
    const source = await client.create('project-one')
    await client.load(source.origin)
    const conversationId = workspaceConversationKey(source.origin)
    await client.entriesReady(true)
    const from = {
      sender: { id: 21, mainFrame: {}, isDestroyed: () => false },
      senderFrame: {},
    } as unknown as IpcMainInvokeEvent
    Object.assign(from, { senderFrame: from.sender.mainFrame })
    f.tabs.tabIdForWebContents.mockReturnValue('A')
    f.tabs.webContentsForTab.mockReturnValue(from.sender)
    const entries: Extract<
      import('../src/shared/workspace-panel-api').WorkspacePanelEvent,
      { kind: 'entry' }
    >[] = []
    client.subscribe((event) => {
      if (event.kind === 'entry') entries.push(event)
    })
    const offered = f.service.offerTemplate(
      from,
      { id: 'project-copy', mode: 'skill', target: 'docs', text: 'Fill', chips: [], createdAt: 1 },
      'project-one',
    )
    await until(() => entries.length === 1)
    expect(entries[0]!.entry).toMatchObject({ projectId: 'project-one' })
    await client.acknowledgeEntry(entries[0]!.entryId!, true, conversationId)
    expect(await offered).toBe(true)
    const other = await client.create()
    await client.load(other.origin)
    await expect(
      client.enqueue({
        conversationId: workspaceConversationKey(other.origin),
        runId: 'wrong-conversation',
        targetId: entries[0]!.capture.id,
        referenceIds: [],
        project: null,
      }),
    ).rejects.toBeDefined()
    const prepared = await client.prepareProject(conversationId, 'project-one')
    const binding = await client.enqueue({
      conversationId,
      runId: 'project-run',
      targetId: entries[0]!.capture.id,
      referenceIds: [],
      project: prepared.binding,
    })
    await client.startNext()
    const original = f.aRequestOriginal
    f.aRequest.mockImplementation((operation, signal) =>
      operation.kind === 'save-project-copy'
        ? (Promise.resolve(true) as never)
        : original(operation, signal),
    )
    await client.complete(binding)
    expect(f.aRequest.mock.calls.filter(([op]) => op.kind === 'save-project-copy')).toEqual([
      [
        expect.objectContaining({ kind: 'save-project-copy', projectId: 'project-one' }),
        expect.any(AbortSignal),
      ],
    ])
    await expect(client.complete(binding)).rejects.toBeDefined()
    expect(f.aRequest.mock.calls.filter(([op]) => op.kind === 'save-project-copy')).toHaveLength(1)
    await client.dispose()
  })
})
