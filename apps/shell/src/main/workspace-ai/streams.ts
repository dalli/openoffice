import type { AiStreamChunk, AiStreamRequest } from '@genoffice/ai-provider'
import type { WorkspaceRunHandle, WorkspaceStartedRun } from './execution-controller'
import { WorkspaceExecutionError } from './execution-controller'
import { randomUUID } from 'node:crypto'

export type WorkspaceStreamRequest = Omit<AiStreamRequest, 'settings'> & { purpose?: 'writer' }
export type WorkspaceStreamService = (
  request: WorkspaceStreamRequest,
  signal: AbortSignal,
  send: (chunk: AiStreamChunk) => void,
) => Promise<void>

const sameRun = (a: WorkspaceRunHandle, b: WorkspaceRunHandle) =>
  a.bindingId === b.bindingId &&
  a.runId === b.runId &&
  a.conversationId === b.conversationId &&
  a.owner.windowId === b.owner.windowId &&
  a.owner.panelContentsId === b.owner.panelContentsId &&
  a.owner.panelGeneration === b.owner.panelGeneration

/** A run owns every model turn until the existing provider has actually stopped. */
export class WorkspaceStreams {
  private current: WorkspaceStartedRun | null = null
  private turn: {
    requestId: string
    abort: AbortController
    done: Promise<void>
  } | null = null
  private closed = false
  private readonly sessions = new Map<string, string>()
  constructor(private readonly service: WorkspaceStreamService) {}

  bind(run: WorkspaceStartedRun) {
    if (
      this.closed ||
      this.turn ||
      run.signal.aborted ||
      (this.current && !this.current.signal.aborted)
    )
      throw new WorkspaceExecutionError('invalid-request')
    this.current = run
  }
  private run(handle: WorkspaceRunHandle) {
    const run = this.current
    if (this.closed || !run || run.signal.aborted || !sameRun(run.binding, handle))
      throw new WorkspaceExecutionError('invalid-request')
    return run
  }
  async stream(
    handle: WorkspaceRunHandle,
    request: WorkspaceStreamRequest,
    send: (chunk: AiStreamChunk) => void,
  ): Promise<void> {
    const run = this.run(handle)
    if (this.turn) throw new WorkspaceExecutionError('invalid-request')
    const abort = new AbortController()
    const stop = () => abort.abort()
    run.signal.addEventListener('abort', stop, { once: true })
    // Install the turn before invoking the service, including synchronous test services.
    const turn = { requestId: request.requestId, abort, done: Promise.resolve() }
    this.turn = turn
    turn.done = Promise.resolve().then(async () => {
      if (abort.signal.aborted) return
      const id = JSON.stringify([run.binding.conversationId, request.purpose ?? 'conversation'])
      if (!this.sessions.has(id)) this.sessions.set(id, randomUUID())
      const { purpose: _purpose, ...payload } = request
      await this.service(
        { ...payload, sessionId: this.sessions.get(id)! },
        abort.signal,
        (chunk) => {
          if (this.turn === turn && !this.closed && chunk.requestId === turn.requestId) send(chunk)
        },
      )
    })
    try {
      await turn.done
    } finally {
      run.signal.removeEventListener('abort', stop)
      if (this.turn === turn) this.turn = null
    }
  }
  async cancel(handle: WorkspaceRunHandle, requestId: string): Promise<void> {
    this.run(handle)
    if (this.turn?.requestId !== requestId) return
    this.turn.abort.abort()
    await this.settled()
  }
  async release(handle: WorkspaceRunHandle): Promise<void> {
    if (!this.current || !sameRun(this.current.binding, handle)) return
    const current = this.current
    this.turn?.abort.abort()
    await this.settled()
    if (this.current === current) this.current = null
  }
  async settled(): Promise<void> {
    await this.turn?.done.catch(() => {})
  }
  dispose() {
    this.closed = true
    this.turn?.abort.abort()
    this.current = null
    this.sessions.clear()
  }
}
