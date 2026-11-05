import { EventEmitter } from 'node:events'
import { mkdtempSync, readFileSync, readdirSync, rmSync } from 'node:fs'
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
  }
  const projects = { listProjects: vi.fn(async (): Promise<ProjectListItem[]> => []) }
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
  const service = registerWorkspacePanelIpc(ipc, (event) =>
    event.sender === (contents as unknown)
      ? {
          window: win as unknown as BrowserWindow,
          tabs: tabs as unknown as TabManager,
          history,
          sources,
          sender,
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
    bRequest,
    contents,
    win,
    tabs,
    api,
    service,
    raw,
    event,
    handlers,
    history,
    liveSources,
    memoryApi,
    sources,
    projects,
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
