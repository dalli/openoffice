import {
  WorkspaceAdapterError,
  type WorkspaceDocumentAdapter,
  type WorkspacePreparedDocument,
} from './workspace-adapter'
import type {
  WorkspaceEditorOperation,
  WorkspaceEditorPort,
  WorkspaceEditorRequest,
  WorkspaceEditorSession,
  WorkspaceEditorValues,
} from './workspace-editor-protocol'

/** Installs an editor-owned executor. No provider, filesystem or public MCP is involved. */
export function hostWorkspaceEditor(
  port: WorkspaceEditorPort,
  adapter: WorkspaceDocumentAdapter,
  /** Existing panel and workspace runs must not edit the same renderer concurrently. */
  isLocallyBusy: () => boolean = () => false,
): () => void {
  let disposed = false
  let session: WorkspacePreparedDocument | null = null
  const pending = new Map<
    string,
    { controller: AbortController; operation: WorkspaceEditorOperation }
  >()
  const seen = new Set<string>()
  const check = (signal: AbortSignal) => {
    if (disposed) throw new WorkspaceAdapterError('disposed')
    if (signal.aborted) throw new WorkspaceAdapterError('cancelled')
  }
  const requireSession = (operation: { sessionId: string; runId: string }) => {
    if (!session || session.id !== operation.sessionId || session.runId !== operation.runId) {
      throw new WorkspaceAdapterError('released')
    }
    return session
  }
  const invoke = async (
    operation: WorkspaceEditorOperation,
    signal: AbortSignal,
  ): Promise<WorkspaceEditorValues[keyof WorkspaceEditorValues]> => {
    check(signal)
    if (operation.kind !== 'release' && isLocallyBusy()) throw new WorkspaceAdapterError('busy')
    switch (operation.kind) {
      case 'capture':
        return adapter.capture(signal)
      case 'read':
        return adapter.read(operation.state, operation.read, signal)
      case 'prepare': {
        const prepared = await adapter.prepare(operation.request, signal)
        try {
          check(signal)
          session = prepared
          return {
            id: prepared.id,
            runId: prepared.runId,
            state: prepared.state,
            systemPrompt: prepared.systemPrompt,
            context: prepared.context,
            tools: prepared.tools,
            toolAccess: Object.fromEntries(
              prepared.tools.map((tool) => [tool.name, adapter.toolAccess[tool.name]]),
            ),
          } satisfies WorkspaceEditorSession
        } catch (error) {
          prepared.release()
          throw error
        }
      }
      case 'execute':
        return requireSession(operation).execute(operation.call, signal)
      case 'verify':
        return requireSession(operation).verifyResponse(operation.text, operation.executed)
      case 'release':
        requireSession(operation).release()
        session = null
        return null
    }
    throw new WorkspaceAdapterError('invalid-call')
  }
  const offRequest = port.onRequest((request: WorkspaceEditorRequest) => {
    if (
      disposed ||
      !request ||
      typeof request.id !== 'string' ||
      !request.id ||
      request.id.length > 256
    )
      return
    // Duplicate IPC delivery must never replay a document edit or a preparation.
    if (seen.has(request.id)) return
    seen.add(request.id)
    if (seen.size > 1024) seen.delete(seen.values().next().value!)
    const controller = new AbortController()
    pending.set(request.id, { controller, operation: request.operation })
    void (async () => {
      try {
        const value = await invoke(request.operation, controller.signal)
        check(controller.signal)
        port.reply({ id: request.id, ok: true, value })
      } catch (error) {
        if (!disposed)
          port.reply({
            id: request.id,
            ok: false,
            error: error instanceof WorkspaceAdapterError ? error.kind : 'failed',
          })
      } finally {
        pending.delete(request.id)
      }
    })()
  })
  const offCancel = port.onCancel((id) => {
    const item = pending.get(id)
    if (!item) return
    item.controller.abort()
    // A timed-out execute/verify must not leave a reusable prepared session.
    if (item.operation.kind === 'execute' || item.operation.kind === 'verify') {
      if (session?.id === item.operation.sessionId && session.runId === item.operation.runId) {
        session.release()
        session = null
      }
    }
  })
  return () => {
    if (disposed) return
    disposed = true
    offRequest()
    offCancel()
    for (const { controller } of pending.values()) controller.abort()
    pending.clear()
    session?.release()
    session = null
    adapter.dispose()
  }
}
