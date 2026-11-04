import { describe, expect, it, vi } from 'vitest'
import {
  createWorkspaceDocumentAdapter,
  type WorkspaceAdapterState,
} from '../src/workspace-adapter'
import { hostWorkspaceEditor } from '../src/workspace-editor-host'
import {
  workspaceEditorPort,
  WORKSPACE_EDITOR_CHANNELS,
  type WorkspaceEditorOperation,
  type WorkspaceEditorReply,
  type WorkspaceEditorRequest,
  type WorkspaceEditorSession,
} from '../src/workspace-editor-protocol'

function fixture() {
  const listeners = new Map<string, (event: unknown, value: unknown) => void>()
  const responses = new Map<string, WorkspaceEditorReply>()
  const ipc = {
    on: vi.fn((channel: string, listener: (event: unknown, value: unknown) => void) => {
      listeners.set(channel, listener)
    }),
    removeListener: vi.fn((channel: string) => {
      listeners.delete(channel)
    }),
    send: vi.fn((channel: string, value: unknown) => {
      if (channel === WORKSPACE_EDITOR_CHANNELS.reply) {
        const reply = value as WorkspaceEditorReply
        responses.set(reply.id, reply)
      }
    }),
  }
  let revision = 1
  const state = (): WorkspaceAdapterState => ({
    instanceId: 'A',
    revision: String(revision),
    selectionId: null,
    readOnly: false,
  })
  const execute = vi.fn(async (_call: unknown, _signal?: AbortSignal) => {
    revision++
    return { output: 'updated', summary: 'write', mutated: true }
  })
  const dispose = vi.fn()
  const changes = new Set<() => void>()
  let busy = false
  const adapter = createWorkspaceDocumentAdapter({
    getState: state,
    read: () => ({ text: '{{private}} UNSAVED', truncated: false }),
    masking: 'applied',
    toolAccess: { write: 'write' },
    subscribeChanges: (handler) => {
      changes.add(handler)
      return () => {
        changes.delete(handler)
      }
    },
    createSkill: () => ({
      id: 'test',
      systemPrompt: 'local',
      tools: [
        { name: 'write', description: 'write', inputSchema: { type: 'object', properties: {} } },
      ],
      executeTool: execute,
    }),
    dispose,
  })
  const stop = hostWorkspaceEditor(workspaceEditorPort(ipc), adapter, () => busy)
  let sequence = 0
  const emit = (request: WorkspaceEditorRequest) =>
    listeners.get(WORKSPACE_EDITOR_CHANNELS.request)?.({}, request)
  const start = (operation: WorkspaceEditorOperation) => {
    const id = `request-${++sequence}`
    emit({ id, operation })
    return id
  }
  const result = async (id: string) => {
    await vi.waitFor(() => expect(responses.has(id)).toBe(true))
    return responses.get(id)!
  }
  const request = async (operation: WorkspaceEditorOperation) => result(start(operation))
  const prepare = async (mode: 'edit' | 'reference' = 'edit') => {
    const reply = await request({ kind: 'prepare', request: { runId: 'P', mode, state: state() } })
    expect(reply.ok).toBe(true)
    return (reply as { value: WorkspaceEditorSession }).value
  }
  return {
    listeners,
    ipc,
    state,
    execute,
    dispose,
    stop,
    start,
    emit,
    result,
    request,
    prepare,
    changed: () => {
      revision++
      for (const notify of changes) notify()
    },
    busy: () => {
      busy = true
    },
  }
}

describe('private workspace editor endpoint', () => {
  it('describes capabilities and publishes invalidation without document text', async () => {
    const host = fixture()
    expect(await host.request({ kind: 'describe' })).toMatchObject({
      ok: true,
      value: {
        state: host.state(),
        toolAccess: { write: 'write' },
        masking: 'applied',
        observesChanges: true,
      },
    })
    host.changed()
    const events = host.ipc.send.mock.calls
      .filter(([channel]) => channel === WORKSPACE_EDITOR_CHANNELS.event)
      .map(([, value]) => value)
    expect(events).toEqual([
      { hostId: expect.any(String), kind: 'ready' },
      { hostId: expect.any(String), kind: 'changed' },
    ])
    expect(JSON.stringify(events)).not.toContain('UNSAVED')
    expect(host.execute).not.toHaveBeenCalled()
    host.stop()
  })

  it('acknowledges settlement only when a cancelled native operation has actually ended', async () => {
    const host = fixture()
    const session = await host.prepare()
    let finish!: () => void
    host.execute.mockImplementationOnce(async () => {
      await new Promise<void>((resolve) => {
        finish = resolve
      })
      return { output: 'ignored', summary: 'write', mutated: true }
    })
    const operation = host.start({
      kind: 'execute',
      sessionId: session.id,
      runId: session.runId,
      call: { id: 'slow', name: 'write', input: {} },
    })
    await vi.waitFor(() => expect(host.execute).toHaveBeenCalledOnce())
    const settled = vi.fn()
    const draining = host.request({ kind: 'settle', runId: session.runId }).then(settled)
    await Promise.resolve()
    expect(settled).not.toHaveBeenCalled()
    finish()
    expect(await host.result(operation)).toMatchObject({ ok: false, error: 'cancelled' })
    await draining
    expect(settled).toHaveBeenCalledWith(expect.objectContaining({ ok: true, value: null }))
    await host.prepare()
    host.stop()
  })
  it('does not apply an old release to a replacement editor using the same run id', async () => {
    const previous = fixture()
    const old = await previous.prepare()
    previous.stop()
    const replacement = fixture()
    const current = await replacement.prepare()
    expect(
      await replacement.request({ kind: 'release', sessionId: old.id, runId: old.runId }),
    ).toMatchObject({ ok: false, error: 'released' })
    expect(
      await replacement.request({
        kind: 'execute',
        sessionId: current.id,
        runId: current.runId,
        call: { id: 'new-call', name: 'write', input: {} },
      }),
    ).toMatchObject({ ok: true })
    replacement.stop()
  })
  it('stays passive until main requests a read and returns serializable prepared capabilities', async () => {
    const host = fixture()
    expect(
      host.ipc.send.mock.calls.filter(([channel]) => channel === WORKSPACE_EDITOR_CHANNELS.reply),
    ).toHaveLength(0)
    expect(host.execute).not.toHaveBeenCalled()
    expect(
      await host.request({ kind: 'read', state: host.state(), read: { scope: 'document' } }),
    ).toMatchObject({
      ok: true,
      value: { text: '{{private}} UNSAVED', masking: 'applied', state: host.state() },
    })
    const session = await host.prepare()
    expect(session.toolAccess).toEqual({ write: 'write' })
    expect(structuredClone(session)).toEqual(session)
    expect('execute' in session).toBe(false)
    host.stop()
  })

  it('requires both the original run and prepared session for tools, verification and release', async () => {
    const host = fixture()
    const session = await host.prepare()
    for (const [sessionId, runId] of [
      ['other', 'P'],
      [session.id, 'Q'],
    ]) {
      for (const operation of [
        { kind: 'execute', sessionId, runId, call: { id: 'write', name: 'write', input: {} } },
        { kind: 'verify', sessionId, runId, text: 'done', executed: [] },
        { kind: 'release', sessionId, runId },
      ] as WorkspaceEditorOperation[]) {
        expect(await host.request(operation)).toMatchObject({ ok: false, error: 'released' })
      }
    }
    expect(host.execute).not.toHaveBeenCalled()
    const operation = {
      kind: 'execute',
      sessionId: session.id,
      runId: 'P',
      call: { id: 'write', name: 'write', input: {} },
    } as const
    const id = host.start(operation)
    expect(await host.result(id)).toMatchObject({ ok: true, value: { state: { revision: '2' } } })
    host.emit({ id, operation })
    await Promise.resolve()
    expect(host.execute).toHaveBeenCalledOnce()
    expect(await host.request(operation)).toMatchObject({ ok: false, error: 'duplicate-call' })
    await host.request({ kind: 'release', sessionId: session.id, runId: 'P' })
    expect(await host.request(operation)).toMatchObject({ ok: false, error: 'released' })
    host.stop()
  })

  it('does not expose writes on a reference session or run beside the existing local panel', async () => {
    const host = fixture()
    const session = await host.prepare('reference')
    expect(session.tools).toEqual([])
    expect(session.toolAccess).toEqual({})
    host.busy()
    expect(await host.request({ kind: 'capture' })).toMatchObject({ ok: false, error: 'busy' })
    expect(
      await host.request({ kind: 'release', sessionId: session.id, runId: 'P' }),
    ).toMatchObject({ ok: true })
    host.stop()
  })

  it('cancels the original operation and drops replies after renderer teardown', async () => {
    const host = fixture()
    const session = await host.prepare()
    let finish!: () => void
    let received: AbortSignal | undefined
    host.execute.mockImplementationOnce(async (_call, signal) => {
      received = signal
      await new Promise<void>((resolve) => {
        finish = resolve
      })
      return { output: 'late', summary: 'write', mutated: true }
    })
    const id = host.start({
      kind: 'execute',
      sessionId: session.id,
      runId: 'P',
      call: { id: 'slow', name: 'write', input: {} },
    })
    await vi.waitFor(() => expect(received).toBeDefined())
    host.listeners.get(WORKSPACE_EDITOR_CHANNELS.cancel)?.({}, id)
    expect(received?.aborted).toBe(true)
    host.stop()
    finish()
    await Promise.resolve()
    expect(host.listeners.size).toBe(0)
    expect(host.dispose).toHaveBeenCalledOnce()
    expect(
      host.ipc.send.mock.calls.some(([, reply]) => (reply as WorkspaceEditorReply).id === id),
    ).toBe(false)
    host.stop()
    expect(host.dispose).toHaveBeenCalledOnce()
  })
})
