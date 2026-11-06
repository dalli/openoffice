import { EventEmitter } from 'node:events'
import type { IpcMain, WebContents } from 'electron'
import { afterEach, describe, expect, it, vi } from 'vitest'
import {
  createWorkspaceDocumentAdapter,
  WorkspaceAdapterError,
  type AgentSkill,
  type AgentToolCall,
  type WorkspaceAdapterState,
} from '@genoffice/agent-core'
import { hostWorkspaceEditor } from '@genoffice/agent-core/workspace-editor-host'
import {
  workspaceEditorPort,
  WORKSPACE_EDITOR_CHANNELS as C,
  type WorkspaceEditorRequest,
} from '@genoffice/agent-core/workspace-editor-protocol'
import { WorkspaceEditorConnection } from '../src/main/workspace-ai/editor-connection'
import { WorkspaceExecutionController } from '../src/main/workspace-ai/execution-controller'

const cleanups: Array<() => void> = []
afterEach(() => {
  cleanups.splice(0).forEach((cleanup) => cleanup())
  vi.useRealTimers()
})
const deferred = () => {
  let resolve!: () => void
  const promise = new Promise<void>((done) => {
    resolve = done
  })
  return { resolve, promise }
}
const tool = (name = 'write', id = 'call-1'): AgentToolCall => ({
  id,
  name,
  input: { text: 'replacement' },
})

function fixture(timeoutMs = 10_000, panelGeneration = 1, writing = false) {
  const ipc = new EventEmitter()
  const owner = { windowId: 1, panelContentsId: 10, panelGeneration }
  const editors = new Map<string, ReturnType<typeof editor>>()
  let sequence = 10
  const stopped = vi.fn()
  const controller = new WorkspaceExecutionController({
    owner,
    timeoutMs,
    onStopped: stopped,
    resolve: (tabId) => {
      const document = editors.get(tabId)
      return document ? { kind: 'docs', client: document.client } : null
    },
  })
  cleanups.push(() => controller.dispose())

  function editor(tabId: string, masking: 'applied' | 'unsupported' = 'applied') {
    let body = `${tabId} UNSAVED`
    let revision = 1
    let instanceId = `${tabId}-instance`
    let readOnly = false
    let selectionId = 'selected-range'
    const changes = new Set<() => void>()
    const renderer = new EventEmitter()
    const view = Object.assign(new EventEmitter(), {
      id: ++sequence,
      mainFrame: {},
      isDestroyed: () => false,
      send: vi.fn((channel: string, value: unknown) => {
        queueMicrotask(() => renderer.emit(channel, {}, structuredClone(value)))
      }),
    })
    const state = (): WorkspaceAdapterState => ({
      instanceId,
      revision: String(revision),
      selectionId,
      readOnly,
    })
    const notify = () => {
      for (const handler of changes) handler()
    }
    const read = vi.fn(() => ({ text: body, truncated: false }))
    const writes = vi.fn()
    const gate = {
      write: null as ReturnType<typeof deferred> | null,
      prepare: null as ReturnType<typeof deferred> | null,
      template: null as ReturnType<typeof deferred> | null,
    }
    const createSkill = vi.fn(async (captured: WorkspaceAdapterState): Promise<AgentSkill> => {
      if (gate.prepare) await gate.prepare.promise
      let expected = captured.revision
      return {
        id: tabId,
        systemPrompt: `instructions-${tabId}`,
        buildContext: () => body,
        tools: ['read', 'write', ...(writing ? ['write_document'] : [])].map((name) => ({
          name,
          description: name,
          inputSchema: { type: 'object', properties: { text: { type: 'string' } } },
        })),
        async executeTool(call, signal) {
          if ((call.name === 'write' || call.name === 'write_document') && gate.write)
            await gate.write.promise
          if (signal?.aborted) throw new WorkspaceAdapterError('cancelled')
          if (expected !== String(revision)) throw new WorkspaceAdapterError('stale-document')
          if (call.name === 'read') return { output: body, summary: 'read' }
          body = String(call.input.text)
          revision++
          expected = String(revision)
          writes(tabId, call)
          notify()
          return { output: 'updated', summary: 'write', mutated: true }
        },
        verifyResponse: (text) => (text === 'incorrect claim' ? 'correction' : null),
      }
    })
    const adapter = createWorkspaceDocumentAdapter({
      masking,
      getState: state,
      read,
      createSkill,
      template: async (_state, mode) => {
        if (gate.template) await gate.template.promise
        return {
          fileName: `${tabId}.docx`,
          title: tabId,
          redacted: false,
          bytes: mode === 'info' ? null : new TextEncoder().encode(body),
        }
      },
      toolAccess: {
        read: 'read',
        write: 'write',
        ...(writing ? { write_document: 'write' as const } : {}),
      },
      subscribeChanges: (handler) => {
        changes.add(handler)
        return () => {
          changes.delete(handler)
        }
      },
    })
    const updateWriter = vi.fn(
      async (_id: string, update: import('@genoffice/agent-core').WorkspaceWriterUpdate) => {
        if (update.kind === 'result') gate.write?.resolve()
      },
    )
    if (writing) adapter.updateWriter = updateWriter
    const saveProjectCopy = vi.fn(async (_id: string, validate: () => Promise<void>) => {
      await validate()
      return true
    })
    const editQueue = vi.fn(async () => ({ version: 'queue-v1', items: [] }))
    const stop = hostWorkspaceEditor(
      workspaceEditorPort({
        on: (channel, listener) => renderer.on(channel, listener),
        removeListener: (channel, listener) => renderer.removeListener(channel, listener),
        send: (channel, value) => {
          queueMicrotask(() =>
            ipc.emit(
              channel,
              { sender: view, senderFrame: view.mainFrame },
              structuredClone(value),
            ),
          )
        },
      }),
      adapter,
      undefined,
      { editQueue, saveProjectCopy, subscribe: () => () => {} },
    )
    const client = new WorkspaceEditorConnection(
      view as unknown as WebContents,
      ipc as IpcMain,
      () => editors.get(tabId)?.client === client,
      1000,
    )
    const document = {
      client,
      editQueue,
      saveProjectCopy,
      updateWriter,
      read,
      writes,
      createSkill,
      gate,
      view,
      state,
      body: () => body,
      manual: (text = 'manual edit') => {
        body = text
        revision++
        notify()
      },
      caret: (id = 'later-caret') => {
        selectionId = id
      },
      lock: () => {
        readOnly = true
        notify()
      },
      reload: () => {
        instanceId += '-replacement'
        notify()
      },
      close: () => {
        client.dispose()
        editors.delete(tabId)
      },
    }
    editors.set(tabId, document)
    cleanups.push(() => {
      stop()
      client.dispose()
    })
    return document
  }
  const a = editor('a')
  const b = editor('b')
  const c = editor('c')
  const capture = async () => {
    const target = await controller.capture('b')
    const reference = await controller.readReference((await controller.capture('a')).id, {
      scope: 'document',
    })
    return { target, reference }
  }
  const enqueue = (
    targetId: string,
    referenceIds: string[] = [],
    runId = 'P',
    conversationId = 'conversation-P',
  ) => controller.enqueue({ targetId, referenceIds, runId, conversationId, project: null })
  return { controller, a, b, c, editor, editors, stopped, capture, enqueue }
}

describe('workspace registry and live editor coordination', () => {
  it('binds edit batches to the exact source version with no visible selection and no body read', async () => {
    const f = fixture()
    const state = f.b.state()
    const expected = { documentId: state.instanceId, revision: state.revision }
    const capture = await f.controller.captureEditBatch('b', expected)
    expect(capture.binding.selection).toBeNull()
    expect(f.b.read).not.toHaveBeenCalled()
    const run = f.enqueue(capture.id)
    expect((await f.controller.startNext())!.target.state.selectionId).toBeNull()
    await f.controller.cancel(run)
    f.b.manual()
    await expect(f.controller.captureEditBatch('b', expected)).rejects.toMatchObject({
      kind: 'document-changed',
    })
    const now = f.b.state()
    f.b.lock()
    await expect(
      f.controller.captureEditBatch('b', { documentId: now.instanceId, revision: now.revision }),
    ).rejects.toMatchObject({ kind: 'document-changed' })
  })
  it('routes writer updates only to the active call and rejects copied owner/conversation handles', async () => {
    const f = fixture(10_000, 1, true)
    const target = await f.controller.capture('b')
    const bound = f.enqueue(target.id)
    await f.controller.startNext()
    await expect(
      f.controller.updateWriter(bound, 'writer', { kind: 'ready' }),
    ).rejects.toMatchObject({ kind: 'invalid-request' })
    f.b.gate.write = deferred()
    const executing = f.controller.execute(
      bound,
      bound.target.document,
      tool('write_document', 'writer'),
    )
    await f.controller.updateWriter(bound, 'writer', { kind: 'ready' })
    for (const handle of [
      { ...bound, conversationId: 'other-conversation' },
      { ...bound, owner: { ...bound.owner, panelGeneration: 2 } },
    ])
      await expect(
        f.controller.updateWriter(handle, 'writer', { kind: 'progress', text: 'wrong' }),
      ).rejects.toThrow()
    await expect(
      f.controller.updateWriter(bound, 'wrong-call', { kind: 'result', ok: true, text: 'wrong' }),
    ).rejects.toMatchObject({ kind: 'invalid-request' })
    expect(f.a.updateWriter).not.toHaveBeenCalled()
    expect(f.c.updateWriter).not.toHaveBeenCalled()
    await f.controller.updateWriter(bound, 'writer', { kind: 'progress', text: 'live' })
    await f.controller.updateWriter(bound, 'writer', { kind: 'result', ok: true, text: 'complete' })
    expect((await executing).result.mutated).toBe(true)
    expect(f.b.updateWriter).toHaveBeenCalledTimes(3)
    await expect(
      f.controller.updateWriter(bound, 'writer', { kind: 'progress', text: 'late' }),
    ).rejects.toMatchObject({ kind: 'invalid-request' })
  })
  it('pins template reads to the original target, rejects replacements and discards mid-export changes', async () => {
    const f = fixture()
    const a = await f.controller.capture('a')
    expect(await f.controller.template(a.binding.document, 'preview')).toMatchObject({
      fileName: 'a.docx',
      bytes: new TextEncoder().encode('a UNSAVED'),
    })
    await expect(
      f.controller.template({ ...a.binding.document, windowId: 2 }, 'preview'),
    ).rejects.toMatchObject({ kind: 'wrong-owner' })
    f.a.gate.template = deferred()
    const reading = f.controller.template(a.binding.document, 'preview')
    await vi.waitFor(() =>
      expect(
        f.a.view.send.mock.calls.some(
          ([, value]) => (value as WorkspaceEditorRequest)?.operation?.kind === 'template',
        ),
      ).toBe(true),
    )
    await new Promise((done) => setTimeout(done, 0))
    f.a.manual()
    f.a.gate.template.resolve()
    await expect(reading).rejects.toMatchObject({ kind: 'document-changed' })
    f.a.reload()
    await expect(f.controller.template(a.binding.document, 'info')).rejects.toMatchObject({
      kind: 'document-changed',
    })
  })
  it('renews the pinned selection without reading content, and rejects a replaced document', async () => {
    const f = fixture()
    const original = await f.controller.capture('b')
    f.b.caret()
    const renewed = await f.controller.refresh(original.id)
    expect(renewed.binding).toEqual(original.binding)
    expect(renewed.id).not.toBe(original.id)
    expect(f.b.read).not.toHaveBeenCalled()
    expect(f.b.createSkill).not.toHaveBeenCalled()
    f.b.reload()
    await expect(f.controller.refresh(renewed.id)).rejects.toMatchObject({
      kind: 'editor-replaced',
    })
    expect(f.b.writes).not.toHaveBeenCalled()
  })

  it('does not renew an already queued binding when a later target capture is refreshed', async () => {
    const f = fixture()
    const pinned = await f.controller.capture('b')
    f.enqueue(pinned.id)
    f.b.manual()
    await vi.waitFor(() => expect(f.controller.state('P')).toMatchObject({ status: 'stopped' }))
    const latest = await f.controller.capture('b')
    const renewed = await f.controller.refresh(latest.id)
    expect(renewed.binding.revision).not.toBe(pinned.binding.revision)
    expect(f.controller.state('P')).toMatchObject({ status: 'stopped', reason: 'document-changed' })
    expect(await f.controller.startNext()).toBeNull()
    expect(f.b.createSkill).not.toHaveBeenCalled()
  })

  it('does not accept an old panel handle when a new generation reuses the run and binding ids', async () => {
    const previous = fixture()
    const old = previous.enqueue((await previous.controller.capture('b')).id)
    await previous.controller.startNext()
    previous.controller.dispose()
    const current = fixture(10_000, 2)
    const run = current.enqueue((await current.controller.capture('b')).id)
    await current.controller.startNext()
    expect(run.bindingId).toBe(old.bindingId)
    await expect(
      current.controller.execute(old, run.target.document, tool()),
    ).rejects.toMatchObject({ kind: 'wrong-owner' })
    expect(() => current.controller.cancel(old)).toThrow('wrong-owner')
    expect(current.controller.state('P')).toEqual({ status: 'running' })
    expect(current.b.writes).not.toHaveBeenCalled()
  })
  it('captures metadata without reading text or preparing a skill', async () => {
    const host = fixture()
    const capture = await host.controller.capture('b')
    expect(capture.binding.document).toMatchObject({
      windowId: 1,
      tabId: 'b',
      documentId: 'b-instance',
    })
    for (const editor of [host.a, host.b, host.c]) {
      expect(editor.read).not.toHaveBeenCalled()
      expect(editor.createSkill).not.toHaveBeenCalled()
    }
    await expect(host.controller.capture('another-window-tab')).rejects.toMatchObject({
      kind: 'missing-document',
    })
  })

  it('uses explicit current reference material and fixes target, selection and conversation at enqueue', async () => {
    const host = fixture()
    const { target, reference } = await host.capture()
    expect(reference.snapshot.text).toBe('a UNSAVED')
    const binding = host.enqueue(target.id, [reference.id])
    Object.assign(target.binding.document, { tabId: 'c' })
    Object.assign(reference.snapshot, { text: 'changed by a different draft' })
    host.b.caret()
    const started = await host.controller.startNext()
    expect(started?.binding).toEqual(binding)
    expect(started?.target.state.selectionId).toBe('selected-range')
    expect(started?.references[0].snapshot.text).toBe('a UNSAVED')
    expect(host.a.createSkill).not.toHaveBeenCalled()
    expect(host.c.createSkill).not.toHaveBeenCalled()
    await host.controller.execute(binding, binding.target.document, tool())
    expect(host.b.body()).toBe('replacement')
    expect(host.a.body()).toBe('a UNSAVED')
    expect(host.c.body()).toBe('c UNSAVED')
    expect(
      await host.controller.verifyResponse(binding, 'incorrect claim', [
        { name: 'write', ok: true },
      ]),
    ).toBe('correction')
    await host.controller.complete(binding)
    expect(host.controller.state('P')).toEqual({ status: 'completed' })
  })

  it('keeps reference scope and masking support and does not expose native reference context', async () => {
    const host = fixture()
    host.editor('unsupported', 'unsupported')
    const reference = await host.controller.readReference(
      (await host.controller.capture('unsupported')).id,
      { scope: 'selection', maxChars: 5 },
    )
    expect(reference.snapshot).toMatchObject({
      scope: 'selection',
      masking: 'unsupported',
      text: 'unsup',
      truncated: true,
    })
    const target = await host.controller.capture('b')
    const binding = host.enqueue(target.id, [reference.id])
    const started = await host.controller.startNext()
    expect(started?.references[0].snapshot).toEqual(reference.snapshot)
    expect(started?.target.context).toBe('b UNSAVED')
    expect(host.editors.get('unsupported')?.createSkill).not.toHaveBeenCalled()
    await expect(
      host.controller.execute(binding, reference.binding.document, tool()),
    ).rejects.toMatchObject({ kind: 'read-only' })
    expect(host.controller.state('P')).toEqual({ status: 'running' })
    host.controller.cancel(binding)
  })

  it('rejects forged snapshot ids, same-document references and another conversation handle', async () => {
    const host = fixture()
    const { target, reference } = await host.capture()
    expect(() => host.enqueue('forged')).toThrow('invalid-request')
    expect(() => host.enqueue(target.id, ['forged'])).toThrow('invalid-request')
    const same = await host.controller.readReference(target.id, { scope: 'document' })
    expect(() => host.enqueue(target.id, [same.id])).toThrow('invalid-request')
    const run = host.enqueue(target.id, [reference.id])
    await host.controller.startNext()
    await expect(
      host.controller.execute({ ...run, conversationId: 'Q' }, run.target.document, tool()),
    ).rejects.toMatchObject({ kind: 'run-not-active' })
    await expect(
      host.controller.execute(run, { ...run.target.document, tabId: 'c' }, tool()),
    ).rejects.toMatchObject({ kind: 'unbound-document' })
    expect(host.b.writes).not.toHaveBeenCalled()
  })

  it('never promotes native reference read tools beyond the explicitly acquired snapshot', async () => {
    const host = fixture()
    const { target, reference } = await host.capture()
    const run = host.enqueue(target.id, [reference.id])
    await host.controller.startNext()
    await expect(
      host.controller.execute(run, reference.binding.document, tool('read')),
    ).rejects.toMatchObject({ kind: 'unknown-tool' })
    expect(host.a.createSkill).not.toHaveBeenCalled()
    expect(host.controller.state('P')).toEqual({ status: 'stopped', reason: 'unknown-tool' })
  })

  it('serializes FIFO and invalidates a queued snapshot after an earlier run edits its target', async () => {
    const host = fixture()
    const target = await host.controller.capture('b')
    const c = await host.controller.capture('c')
    const p = host.enqueue(target.id)
    host.enqueue(target.id, [], 'Q', 'conversation-Q')
    const r = host.enqueue(c.id, [], 'R', 'conversation-R')
    await host.controller.startNext()
    expect(await host.controller.startNext()).toBeNull()
    await host.controller.execute(p, p.target.document, tool())
    expect(host.controller.state('Q')).toEqual({ status: 'stopped', reason: 'document-changed' })
    await host.controller.complete(p)
    expect((await host.controller.startNext())?.binding).toEqual(r)
  })

  it('rejects an edited or reloaded queued document before creating any skill', async () => {
    const host = fixture()
    const target = await host.controller.capture('b')
    host.enqueue(target.id)
    host.b.reload()
    await Promise.resolve()
    expect(await host.controller.startNext()).toBeNull()
    expect(host.b.createSkill).not.toHaveBeenCalled()
    expect(host.controller.state('P')?.status).toBe('stopped')
    expect(() => host.enqueue(target.id, [], 'Q')).toThrow('document-changed')
  })

  it('aborts a delayed target operation when its reference changes and waits for editor idleness', async () => {
    const host = fixture()
    const { target, reference } = await host.capture()
    const p = host.enqueue(target.id, [reference.id])
    host.enqueue((await host.controller.capture('c')).id, [], 'Q', 'conversation-Q')
    const started = await host.controller.startNext()
    host.b.gate.write = deferred()
    const pending = host.controller.execute(p, p.target.document, tool())
    const rejected = expect(pending).rejects.toBeDefined()
    await vi.waitFor(() =>
      expect(
        host.b.view.send.mock.calls.some(
          ([channel, value]) =>
            channel === C.request && (value as WorkspaceEditorRequest).operation.kind === 'execute',
        ),
      ).toBe(true),
    )
    host.a.manual('new reference')
    await vi.waitFor(() => expect(started?.signal.aborted).toBe(true))
    expect(host.controller.state('P')).toEqual({ status: 'stopped', reason: 'document-changed' })
    const next = host.controller.startNext()
    await Promise.resolve()
    expect(host.c.createSkill).not.toHaveBeenCalled()
    host.b.gate.write.resolve()
    await rejected
    expect((await next)?.binding.runId).toBe('Q')
    expect(host.b.writes).not.toHaveBeenCalled()
  })

  it('settles cancelled preparation before allowing the next run on that editor', async () => {
    const host = fixture()
    const target = await host.controller.capture('b')
    const p = host.enqueue(target.id)
    host.enqueue(target.id, [], 'Q', 'conversation-Q')
    host.b.gate.prepare = deferred()
    const preparing = host.controller.startNext()
    const rejected = expect(preparing).rejects.toBeDefined()
    await vi.waitFor(() => expect(host.b.createSkill).toHaveBeenCalledOnce())
    host.controller.cancel(p)
    const next = host.controller.startNext()
    host.b.gate.prepare.resolve()
    await rejected
    // startNext calls made during preparation do not consume the next FIFO entry.
    expect(await next).toBeNull()
    expect((await host.controller.startNext())?.binding.runId).toBe('Q')
  })

  it('rejects duplicate calls and another simultaneous tool without replaying an edit', async () => {
    const host = fixture()
    const run = host.enqueue((await host.controller.capture('b')).id)
    await host.controller.startNext()
    host.b.gate.write = deferred()
    const first = host.controller.execute(run, run.target.document, tool())
    await expect(
      host.controller.execute(run, run.target.document, tool('write', 'other')),
    ).rejects.toMatchObject({ kind: 'call-in-flight' })
    host.b.gate.write.resolve()
    await first
    await expect(host.controller.execute(run, run.target.document, tool())).rejects.toMatchObject({
      kind: 'duplicate-call',
    })
    expect(host.b.writes).toHaveBeenCalledOnce()
  })

  it('expires idle runs via the timer and invalidates old handles when the panel is disposed', async () => {
    vi.useFakeTimers()
    const host = fixture(100)
    const run = host.enqueue((await host.controller.capture('b')).id)
    const started = await host.controller.startNext()
    await vi.advanceTimersByTimeAsync(100)
    expect(started?.signal.aborted).toBe(true)
    expect(host.controller.state('P')).toEqual({ status: 'stopped', reason: 'timeout' })
    host.controller.dispose()
    await expect(host.controller.execute(run, run.target.document, tool())).rejects.toMatchObject({
      kind: 'panel-closed',
    })
  })

  it('cancels a run when a native editor is closed without touching another conversation', async () => {
    const host = fixture()
    const { target, reference } = await host.capture()
    const p = host.enqueue(target.id, [reference.id])
    const q = host.enqueue((await host.controller.capture('c')).id, [], 'Q', 'conversation-Q')
    const started = await host.controller.startNext()
    host.a.close()
    expect(started?.signal.aborted).toBe(true)
    expect(host.controller.state('P')).toEqual({ status: 'stopped', reason: 'document-closed' })
    expect((await host.controller.startNext())?.binding).toEqual(q)
    await expect(host.controller.execute(p, p.target.document, tool())).rejects.toMatchObject({
      kind: 'run-not-active',
    })
    expect(host.b.writes).not.toHaveBeenCalled()
  })
})

describe('main-owned queue controls', () => {
  it('routes only to the original document and rejects other windows, generations and active runs', async () => {
    const f = fixture()
    const target = await f.controller.capture('b')
    const doc = target.binding.document
    expect(await f.controller.editQueue(doc, { kind: 'list' })).toEqual({
      version: 'queue-v1',
      items: [],
    })
    expect(f.b.editQueue).toHaveBeenCalledOnce()
    expect(f.a.editQueue).not.toHaveBeenCalled()
    await expect(
      f.controller.editQueue({ ...doc, windowId: 2 }, { kind: 'list' }),
    ).rejects.toMatchObject({ kind: 'wrong-owner' })
    await expect(
      f.controller.editQueue({ ...doc, generation: 999 }, { kind: 'clear', version: 'queue-v1' }),
    ).rejects.toMatchObject({ kind: 'editor-replaced' })
    const fresh = await f.controller.capture('b')
    f.controller.enqueue({
      conversationId: 'P',
      runId: 'run',
      targetId: fresh.id,
      referenceIds: [],
      project: null,
    })
    await f.controller.startNext()
    await expect(
      f.controller.editQueue(doc, { kind: 'send', version: 'queue-v1' }),
    ).rejects.toMatchObject({ kind: 'call-in-flight' })
    expect(f.b.editQueue).toHaveBeenCalledOnce()
  })
})

describe('queue snapshot lifetime', () => {
  it('keeps the original conversation capture after many card refreshes', async () => {
    const f = fixture()
    const capture = await f.controller.capture('b')
    for (let i = 0; i < 140; i++)
      await f.controller.editQueue(capture.binding.document, { kind: 'list' })
    expect(await f.controller.refresh(capture.id)).toMatchObject({ binding: capture.binding })
  })
})

describe('project copy completion', () => {
  it('saves only after a matching successful run is settled and never completes twice', async () => {
    const f = fixture()
    const target = await f.controller.capture('b')
    const run = f.controller.enqueue({
      conversationId: 'project-chat',
      runId: 'project-run',
      targetId: target.id,
      referenceIds: [],
      project: { id: 'project-one', contextVersion: 'context-one' },
    })
    await f.controller.startNext()
    await f.controller.execute(run, run.target.document, tool())
    await expect(f.controller.completeProjectCopy(run, 'other-project')).rejects.toMatchObject({
      kind: 'invalid-request',
    })
    expect(f.b.saveProjectCopy).not.toHaveBeenCalled()
    await f.controller.completeProjectCopy(run, 'project-one')
    expect(f.b.saveProjectCopy).toHaveBeenCalledOnce()
    expect(f.b.saveProjectCopy).toHaveBeenCalledWith('project-one', expect.any(Function))
    await expect(f.controller.completeProjectCopy(run, 'project-one')).rejects.toBeDefined()
    expect(f.b.saveProjectCopy).toHaveBeenCalledOnce()
  })
  it.each(['cancel', 'change'] as const)('does not save a run after %s', async (reason) => {
    const f = fixture()
    const target = await f.controller.capture('b')
    const run = f.controller.enqueue({
      conversationId: 'project-chat',
      runId: 'project-run',
      targetId: target.id,
      referenceIds: [],
      project: { id: 'project-one', contextVersion: 'context-one' },
    })
    await f.controller.startNext()
    if (reason === 'cancel') f.controller.cancel(run)
    else f.b.manual()
    await expect(f.controller.completeProjectCopy(run, 'project-one')).rejects.toBeDefined()
    expect(f.b.saveProjectCopy).not.toHaveBeenCalled()
  })
})

it('holds the next workspace run until a project copy save settles and cancels its late write', async () => {
  const f = fixture()
  const target = await f.controller.capture('b')
  const later = await f.controller.capture('c')
  const run = f.controller.enqueue({
    conversationId: 'copy-chat',
    runId: 'copy-run',
    targetId: target.id,
    referenceIds: [],
    project: { id: 'project-one', contextVersion: 'context-one' },
  })
  await f.controller.startNext()
  const waiting = deferred()
  const written = vi.fn()
  f.b.saveProjectCopy.mockImplementationOnce(async (_id, validate) => {
    await waiting.promise
    await validate()
    written()
    return true
  })
  const finishing = f.controller.completeProjectCopy(run, 'project-one')
  const result = expect(finishing).rejects.toMatchObject({ kind: 'cancelled' })
  await vi.waitFor(() => expect(f.b.saveProjectCopy).toHaveBeenCalledOnce())
  f.controller.enqueue({
    conversationId: 'later-chat',
    runId: 'later-run',
    targetId: later.id,
    referenceIds: [],
    project: null,
  })
  expect(await f.controller.startNext()).toBeNull()
  f.controller.cancel(run)
  let settled = false
  void finishing
    .catch(() => {})
    .finally(() => {
      settled = true
    })
  await new Promise<void>((resolve) => setTimeout(resolve, 0))
  expect(settled).toBe(false)
  expect(await f.controller.startNext()).toBeNull()
  waiting.resolve()
  await result
  await new Promise<void>((resolve) => setTimeout(resolve, 0))
  expect(written).not.toHaveBeenCalled()
  expect((await f.controller.startNext())?.binding.runId).toBe('later-run')
})
