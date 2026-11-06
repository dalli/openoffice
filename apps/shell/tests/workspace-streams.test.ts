import { describe, expect, it, vi } from 'vitest'
import { WorkspaceStreams, type WorkspaceStreamRequest } from '../src/main/workspace-ai/streams'
import type { WorkspaceStartedRun } from '../src/main/workspace-ai/execution-controller'
import { deferred } from './helpers/workspace-host-fixture'

function run(id = 1) {
  const abort = new AbortController()
  return {
    abort,
    value: {
      binding: {
        bindingId: id,
        runId: `run-${id}`,
        conversationId: `conversation-${id}`,
        owner: { windowId: 1, panelContentsId: 10, panelGeneration: 1 },
      },
      signal: abort.signal,
    } as WorkspaceStartedRun,
  }
}
const request: WorkspaceStreamRequest = {
  requestId: 'request-one',
  sessionId: 'session-one',
  system: '',
  messages: [],
  tools: [],
}

describe('workspace model turns', () => {
  it('keeps the tool-less writer in a separate main-issued native session', async () => {
    const service = vi.fn(async () => {})
    const streams = new WorkspaceStreams(service)
    const a = run()
    streams.bind(a.value)
    await streams.stream(a.value.binding, request, vi.fn())
    await streams.stream(a.value.binding, { ...request, purpose: 'writer' }, vi.fn())
    await streams.stream(a.value.binding, request, vi.fn())
    const turns = service.mock.calls as unknown as [WorkspaceStreamRequest][]
    expect(turns[0]![0].sessionId).toBe(turns[2]![0].sessionId)
    expect(turns[1]![0].sessionId).not.toBe(turns[0]![0].sessionId)
    expect(turns[1]![0]).not.toHaveProperty('purpose')
    streams.dispose()
  })
  it('assigns distinct native sessions to conversations, keeping each stable across runs', async () => {
    const service = vi.fn(async () => {})
    const streams = new WorkspaceStreams(service)
    const a = run()
    const b = run(2)
    for (const current of [a, b, a]) {
      streams.bind(current.value)
      await streams.stream(current.value.binding, request, vi.fn())
      await streams.release(current.value.binding)
    }
    const ids = service.mock.calls.map(
      (args) => (args as unknown as [WorkspaceStreamRequest])[0].sessionId,
    )
    expect(ids[0]).toBe(ids[2])
    expect(ids[0]).not.toBe(ids[1])
    expect(ids).not.toContain(request.sessionId)
  })
  it('rejects unbound, stale and cross-conversation requests without calling a provider', async () => {
    const service = vi.fn(async () => {})
    const streams = new WorkspaceStreams(service)
    const a = run()
    const send = vi.fn()
    await expect(streams.stream(a.value.binding, request, send)).rejects.toThrow()
    streams.bind(a.value)
    for (const binding of [
      { ...a.value.binding, conversationId: 'other' },
      { ...a.value.binding, bindingId: 2 },
      { ...a.value.binding, owner: { ...a.value.binding.owner, panelGeneration: 2 } },
    ])
      await expect(streams.stream(binding, request, send)).rejects.toThrow()
    expect(service).not.toHaveBeenCalled()
  })

  it('waits for actual provider cancellation before releasing a run or binding the next', async () => {
    const stopped = deferred()
    let signal: AbortSignal | undefined
    const service = vi.fn(async (_request, abort) => {
      signal = abort
      await stopped.promise
    })
    const streams = new WorkspaceStreams(service)
    const a = run()
    streams.bind(a.value)
    const sending = streams.stream(a.value.binding, request, vi.fn())
    await Promise.resolve()
    await expect(streams.stream(a.value.binding, request, vi.fn())).rejects.toThrow()
    const released = vi.fn()
    const release = streams.release(a.value.binding).then(released)
    expect(signal?.aborted).toBe(true)
    expect(() => streams.bind(run(2).value)).toThrow()
    await Promise.resolve()
    expect(released).not.toHaveBeenCalled()
    stopped.resolve()
    await Promise.all([sending, release])
    streams.bind(run(2).value)
  })

  it('ties cancellation and chunks to the current run and request, and stops on disposal', async () => {
    const stopped = deferred()
    let publish: Parameters<ConstructorParameters<typeof WorkspaceStreams>[0]>[2] | undefined
    let signal: AbortSignal | undefined
    const streams = new WorkspaceStreams(async (_request, abort, send) => {
      publish = send
      signal = abort
      await stopped.promise
    })
    const a = run()
    streams.bind(a.value)
    const send = vi.fn()
    const sending = streams.stream(a.value.binding, request, send)
    await Promise.resolve()
    await streams.cancel(a.value.binding, 'other-request')
    await streams.release(run(2).value.binding)
    expect(signal?.aborted).toBe(false)
    publish?.({ requestId: 'other-request', type: 'delta', text: 'WRONG' })
    publish?.({ requestId: request.requestId, type: 'delta', text: 'RIGHT' })
    expect(send).toHaveBeenCalledOnce()
    a.abort.abort()
    expect(signal?.aborted).toBe(true)
    streams.dispose()
    publish?.({ requestId: request.requestId, type: 'done' })
    expect(send).toHaveBeenCalledOnce()
    stopped.resolve()
    await sending
    await expect(streams.stream(a.value.binding, request, send)).rejects.toThrow()
  })
})
