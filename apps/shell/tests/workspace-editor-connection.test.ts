import { EventEmitter } from 'node:events'
import type { IpcMain, WebContents } from 'electron'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { WorkspaceEditorConnection } from '../src/main/workspace-ai/editor-connection'
import {
  WORKSPACE_EDITOR_CHANNELS as C,
  type WorkspaceEditorRequest,
} from '@genoffice/agent-core/workspace-editor-protocol'

const connections: WorkspaceEditorConnection[] = []
afterEach(() => {
  connections.splice(0).forEach((connection) => connection.dispose())
  vi.useRealTimers()
})

function fixture() {
  const ipc = new EventEmitter()
  const view = Object.assign(new EventEmitter(), {
    id: 42,
    mainFrame: {},
    isDestroyed: vi.fn(() => false),
    send: vi.fn(),
  })
  let owned = true
  const connection = new WorkspaceEditorConnection(
    view as unknown as WebContents,
    ipc as IpcMain,
    () => owned,
    100,
  )
  connections.push(connection)
  const sent = (): WorkspaceEditorRequest =>
    view.send.mock.calls.filter(([channel]) => channel === C.request).at(-1)![1]
  const reply = (
    id: string,
    value: unknown,
    sender: unknown = view,
    frame: unknown = view.mainFrame,
  ) => ipc.emit(C.reply, { sender, senderFrame: frame }, { id, ok: true, value })
  return {
    ipc,
    view,
    connection,
    sent,
    reply,
    unown: () => {
      owned = false
    },
  }
}

describe('main-owned workspace editor connection', () => {
  it('accepts metadata changes only from its main frame and invalidates a replaced renderer host', async () => {
    const host = fixture()
    const changed = vi.fn()
    const closed = vi.fn()
    host.connection.onChanged(changed)
    host.connection.onClosed(closed)
    const state = { instanceId: 'A', revision: '1', selectionId: null, readOnly: false }
    const pending = host.connection.request({ kind: 'describe' })
    host.reply(host.sent().id, {
      hostId: 'host-1',
      state,
      toolAccess: {},
      observesChanges: true,
      masking: 'applied',
    })
    await pending
    host.ipc.emit(
      C.event,
      { sender: host.view, senderFrame: {} },
      { hostId: 'host-1', kind: 'changed' },
    )
    expect(changed).not.toHaveBeenCalled()
    host.ipc.emit(
      C.event,
      { sender: host.view, senderFrame: host.view.mainFrame },
      { hostId: 'host-1', kind: 'changed' },
    )
    expect(changed).toHaveBeenCalledOnce()
    host.ipc.emit(
      C.event,
      { sender: host.view, senderFrame: host.view.mainFrame },
      { hostId: 'host-2', kind: 'ready' },
    )
    expect(host.connection.available).toBe(false)
    expect(closed).toHaveBeenCalledOnce()
    expect(host.ipc.listenerCount(C.event)).toBe(0)
  })
  it('rejects malformed replies even when the request id and sender match', async () => {
    const host = fixture()
    const pending = host.connection.request({ kind: 'capture' })
    host.reply(host.sent().id, { instanceId: 'A', revision: 1 })
    await expect(pending).rejects.toMatchObject({ kind: 'wrong-response' })
    expect(host.connection.available).toBe(false)
  })

  it('rejects a prepared session for a different run', async () => {
    const host = fixture()
    const state = { instanceId: 'A', revision: '1', selectionId: null, readOnly: false }
    const pending = host.connection.request({
      kind: 'prepare',
      request: { runId: 'P', state, mode: 'edit' },
    })
    host.reply(host.sent().id, {
      id: 'session',
      runId: 'Q',
      state,
      systemPrompt: '',
      context: '',
      tools: [],
      toolAccess: {},
    })
    await expect(pending).rejects.toMatchObject({ kind: 'wrong-response' })
  })
  it('correlates replies with the live view, main frame and request id', async () => {
    const host = fixture()
    const settled = vi.fn()
    const pending = host.connection.request({ kind: 'capture' }).then(settled)
    const { id } = host.sent()
    host.reply('another-id', { revision: 'wrong' })
    host.reply(id, { revision: 'wrong' }, { id: 42 })
    host.reply(id, { revision: 'wrong' }, host.view, {})
    await Promise.resolve()
    expect(settled).not.toHaveBeenCalled()
    const state = { instanceId: 'A', revision: '1', selectionId: null, readOnly: false }
    host.reply(id, state)
    await pending
    expect(settled).toHaveBeenCalledWith(state)
  })

  it('copies a queued operation and forwards structured adapter failures', async () => {
    const host = fixture()
    const operation = {
      kind: 'read',
      state: { instanceId: 'A', revision: '1', selectionId: null, readOnly: false },
      read: { scope: 'document' },
    } as const
    const pending = host.connection.request(operation)
    const { id } = host.sent()
    Object.assign(operation.state, { instanceId: 'B' })
    expect(host.sent().operation).toMatchObject({ state: { instanceId: 'A' } })
    host.ipc.emit(
      C.reply,
      { sender: host.view, senderFrame: host.view.mainFrame },
      { id, ok: false, error: 'stale-document' },
    )
    await expect(pending).rejects.toMatchObject({ kind: 'stale-document' })
  })

  it.each(['timeout', 'cancelled'] as const)(
    'cancels %s in the editor and never resends the edit',
    async (reason) => {
      vi.useFakeTimers()
      const host = fixture()
      const controller = new AbortController()
      const pending = host.connection.request({ kind: 'capture' }, controller.signal)
      const rejected = expect(pending).rejects.toMatchObject({ kind: reason })
      const { id } = host.sent()
      if (reason === 'timeout') await vi.advanceTimersByTimeAsync(100)
      else controller.abort()
      await rejected
      expect(host.view.send).toHaveBeenCalledWith(C.cancel, id)
      expect(host.view.send.mock.calls.filter(([channel]) => channel === C.request)).toHaveLength(1)
      host.reply(id, { instanceId: 'late' })
      expect(host.connection.available).toBe(reason === 'cancelled')
      expect(host.ipc.listenerCount(C.reply)).toBe(reason === 'cancelled' ? 1 : 0)
    },
  )

  it.each(['destroyed', 'render-process-gone', 'navigation', 'detached'] as const)(
    'invalidates pending work on %s',
    async (event) => {
      const host = fixture()
      const pending = host.connection.request({ kind: 'capture' })
      const rejected = expect(pending).rejects.toMatchObject({ kind: 'closed' })
      if (event === 'navigation') host.view.emit('did-start-navigation', {}, 'reload', false, true)
      else if (event === 'detached') {
        host.unown()
        host.reply(host.sent().id, {})
      } else host.view.emit(event)
      await rejected
      expect(host.connection.available).toBe(false)
      expect(host.ipc.listenerCount(C.reply)).toBe(0)
    },
  )

  it('retains a connection for same-document and subframe navigation', async () => {
    const host = fixture()
    host.view.emit('did-start-navigation', {}, '#section', true, true)
    host.view.emit('did-start-navigation', {}, 'iframe', false, false)
    expect(host.connection.available).toBe(true)
  })

  it('releases an idle prepared session when its owning connection closes', async () => {
    const host = fixture()
    const state = { instanceId: 'A', revision: '1', selectionId: null, readOnly: false }
    const pending = host.connection.request({
      kind: 'prepare',
      request: { runId: 'P', state, mode: 'edit' },
    })
    host.reply(host.sent().id, {
      id: 'session-1',
      runId: 'P',
      state,
      systemPrompt: '',
      context: '',
      tools: [],
      toolAccess: {},
    })
    await pending
    host.connection.dispose()
    expect(host.sent().operation).toEqual({ kind: 'release', sessionId: 'session-1', runId: 'P' })
    expect(host.view.listenerCount('destroyed')).toBe(0)
  })
})
