import { randomUUID } from 'node:crypto'
import type { IpcMain, IpcMainEvent, WebContents } from 'electron'
import type { WorkspaceAdapterState } from '@genoffice/agent-core'
import {
  WORKSPACE_EDITOR_CHANNELS,
  type WorkspaceEditorOperation,
  type WorkspaceEditorReply,
  type WorkspaceEditorSession,
  type WorkspaceEditorValues,
} from '@genoffice/agent-core/workspace-editor-protocol'

interface Pending {
  resolve(value: WorkspaceEditorValues[keyof WorkspaceEditorValues]): void
  reject(error: Error): void
  cleanup(): void
  readonly operation: WorkspaceEditorOperation
}

const record = (value: unknown): value is Record<string, unknown> =>
  !!value && typeof value === 'object' && !Array.isArray(value)
const text = (value: unknown): value is string => typeof value === 'string' && value.length > 0
const stateOf = (value: unknown): value is WorkspaceAdapterState =>
  record(value) &&
  text(value.instanceId) &&
  text(value.revision) &&
  (value.selectionId === null || text(value.selectionId)) &&
  typeof value.readOnly === 'boolean'
const sameState = (left: WorkspaceAdapterState, right: WorkspaceAdapterState) =>
  left.instanceId === right.instanceId &&
  left.revision === right.revision &&
  left.selectionId === right.selectionId &&
  left.readOnly === right.readOnly

function validValue(operation: WorkspaceEditorOperation, value: unknown): boolean {
  switch (operation.kind) {
    case 'capture':
      return stateOf(value)
    case 'read':
      return (
        record(value) &&
        stateOf(value.state) &&
        sameState(value.state, operation.state) &&
        typeof value.text === 'string' &&
        typeof value.truncated === 'boolean' &&
        value.scope === operation.read.scope &&
        (value.masking === 'applied' || value.masking === 'unsupported')
      )
    case 'prepare': {
      if (
        !record(value) ||
        !text(value.id) ||
        value.runId !== operation.request.runId ||
        !stateOf(value.state) ||
        !sameState(value.state, operation.request.state) ||
        typeof value.systemPrompt !== 'string' ||
        typeof value.context !== 'string' ||
        !Array.isArray(value.tools) ||
        !record(value.toolAccess)
      )
        return false
      const access = value.toolAccess
      return (
        value.tools.every(
          (tool) =>
            record(tool) &&
            text(tool.name) &&
            typeof tool.description === 'string' &&
            record(tool.inputSchema) &&
            (access[tool.name] === 'read' ||
              (operation.request.mode === 'edit' &&
                !operation.request.state.readOnly &&
                access[tool.name] === 'write')),
        ) && new Set(value.tools.map((tool) => tool.name)).size === value.tools.length
      )
    }
    case 'execute':
      return (
        record(value) &&
        stateOf(value.state) &&
        record(value.result) &&
        typeof value.result.output === 'string' &&
        typeof value.result.summary === 'string' &&
        (value.result.mutated === undefined || typeof value.result.mutated === 'boolean')
      )
    case 'verify':
      return value === null || typeof value === 'string'
    case 'release':
      return value === null
  }
}

export class WorkspaceEditorConnectionError extends Error {
  constructor(readonly kind: string) {
    super(`Workspace editor connection rejected: ${kind}`)
    this.name = 'WorkspaceEditorConnectionError'
  }
}

/** Main owns the live view and its lifetime. No renderer may supply a destination id. */
export class WorkspaceEditorConnection {
  private readonly prefix = randomUUID()
  private sequence = 0
  private disposed = false
  private readonly pending = new Map<string, Pending>()
  private session: Pick<WorkspaceEditorSession, 'id' | 'runId'> | null = null

  constructor(
    private readonly contents: WebContents,
    private readonly ipc: Pick<IpcMain, 'on' | 'removeListener'>,
    private readonly isOwned: () => boolean,
    private readonly timeoutMs = 30_000,
  ) {
    if (!Number.isFinite(timeoutMs) || timeoutMs <= 0) throw new Error('invalid editor timeout')
    ipc.on(WORKSPACE_EDITOR_CHANNELS.reply, this.onReply)
    contents.on('destroyed', this.onGone)
    contents.on('render-process-gone', this.onGone)
    contents.on('did-start-navigation', this.onNavigation)
  }

  get available(): boolean {
    return !this.disposed && !this.contents.isDestroyed() && this.isOwned()
  }

  request<O extends WorkspaceEditorOperation>(
    operation: O,
    signal?: AbortSignal,
  ): Promise<WorkspaceEditorValues[O['kind']]> {
    if (!this.available) {
      this.dispose()
      return Promise.reject(new WorkspaceEditorConnectionError('unavailable'))
    }
    if (signal?.aborted) return Promise.reject(new WorkspaceEditorConnectionError('cancelled'))
    const id = `${this.prefix}:${++this.sequence}`
    // Clone before yielding; a caller changing its draft cannot retarget a sent operation.
    const message = { id, operation: structuredClone(operation) }
    return new Promise((resolve, reject) => {
      const cancel = (kind: 'cancelled' | 'timeout') => {
        const pending = this.pending.get(id)
        if (!pending) return
        this.pending.delete(id)
        pending.cleanup()
        this.sendCancel(id)
        pending.reject(new WorkspaceEditorConnectionError(kind))
        // A timed-out edit is never retried on this connection.
        this.dispose()
      }
      const abort = () => cancel('cancelled')
      const timer = setTimeout(() => cancel('timeout'), this.timeoutMs)
      this.pending.set(id, {
        operation: message.operation,
        resolve: (value) => resolve(value as WorkspaceEditorValues[O['kind']]),
        reject,
        cleanup: () => {
          clearTimeout(timer)
          signal?.removeEventListener('abort', abort)
        },
      })
      signal?.addEventListener('abort', abort, { once: true })
      try {
        this.contents.send(WORKSPACE_EDITOR_CHANNELS.request, message)
      } catch {
        this.dispose()
      }
    })
  }

  dispose(): void {
    if (this.disposed) return
    this.disposed = true
    for (const [id, pending] of this.pending) {
      pending.cleanup()
      this.sendCancel(id)
      pending.reject(new WorkspaceEditorConnectionError('closed'))
    }
    this.pending.clear()
    if (this.session && !this.contents.isDestroyed()) {
      try {
        this.contents.send(WORKSPACE_EDITOR_CHANNELS.request, {
          id: `${this.prefix}:${++this.sequence}`,
          operation: { kind: 'release', sessionId: this.session.id, runId: this.session.runId },
        })
      } catch {
        /* The view may have closed between the check and the send. */
      }
    }
    this.session = null
    this.ipc.removeListener(WORKSPACE_EDITOR_CHANNELS.reply, this.onReply)
    this.contents.removeListener('destroyed', this.onGone)
    this.contents.removeListener('render-process-gone', this.onGone)
    this.contents.removeListener('did-start-navigation', this.onNavigation)
  }

  private readonly onReply = (event: IpcMainEvent, reply: WorkspaceEditorReply): void => {
    // Subframes and another document with a copied request id cannot answer.
    if (event.sender !== this.contents || event.senderFrame !== this.contents.mainFrame) return
    if (!this.available) {
      this.dispose()
      return
    }
    if (!reply || typeof reply.id !== 'string') return
    const pending = this.pending.get(reply.id)
    if (!pending || typeof reply.ok !== 'boolean') return
    this.pending.delete(reply.id)
    pending.cleanup()
    if (!reply.ok) {
      pending.reject(
        new WorkspaceEditorConnectionError(
          typeof reply.error === 'string' ? reply.error : 'failed',
        ),
      )
      return
    }
    if (!validValue(pending.operation, reply.value)) {
      pending.reject(new WorkspaceEditorConnectionError('wrong-response'))
      this.dispose()
      return
    }
    if (pending.operation.kind === 'prepare') {
      const prepared = reply.value as WorkspaceEditorSession | null
      if (!prepared || typeof prepared.id !== 'string' || typeof prepared.runId !== 'string') {
        pending.reject(new WorkspaceEditorConnectionError('wrong-response'))
        this.dispose()
        return
      }
      this.session = { id: prepared.id, runId: prepared.runId }
    } else if (pending.operation.kind === 'release') this.session = null
    pending.resolve(reply.value)
  }

  private readonly onGone = (): void => {
    this.dispose()
  }
  private readonly onNavigation = (
    _event: unknown,
    _url: string,
    isInPlace: boolean,
    isMainFrame: boolean,
  ): void => {
    if (isMainFrame && !isInPlace) this.dispose()
  }
  private sendCancel(id: string): void {
    if (this.contents.isDestroyed()) return
    try {
      this.contents.send(WORKSPACE_EDITOR_CHANNELS.cancel, id)
    } catch {
      /* Renderer loss already invalidates the request. */
    }
  }
}
