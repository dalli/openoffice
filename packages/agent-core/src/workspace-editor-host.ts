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
  const hostId = crypto.randomUUID()
  let session: WorkspacePreparedDocument | null = null
  const pending = new Map<
    string,
    {
      controller: AbortController
      operation: WorkspaceEditorOperation
      done: Promise<void>
      settle(): void
    }
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
    if (operation.kind !== 'release' && operation.kind !== 'settle' && isLocallyBusy())
      throw new WorkspaceAdapterError('busy')
    switch (operation.kind) {
      case 'settle': {
        const operations = [...pending.values()].filter(({ operation: item }) =>
          item.kind === 'prepare'
            ? item.request.runId === operation.runId
            : (item.kind === 'execute' || item.kind === 'verify') && item.runId === operation.runId,
        )
        for (const item of operations) item.controller.abort()
        if (session?.runId === operation.runId) session.release()
        await Promise.all(operations.map((item) => item.done))
        if (session?.runId === operation.runId) {
          session.release()
          session = null
        }
        return null
      }
      case 'describe':
        return {
          hostId,
          state: await adapter.capture(signal),
          toolAccess: adapter.toolAccess,
          masking: adapter.masking,
          observesChanges: adapter.observesChanges,
        }
      case 'capture':
        return adapter.capture(signal)
      case 'rollback':
        if (!adapter.rollback) throw new WorkspaceAdapterError('released')
        return adapter.rollback(operation.rollbackId, operation.state, signal)
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
      case 'execute': {
        const prepared = requireSession(operation)
        const result = await prepared.execute(operation.call, signal)
        return { ...result, ...(prepared.rollbackId ? { rollbackId: prepared.rollbackId } : {}) }
      }
      case 'verify':
        return requireSession(operation).verifyResponse(operation.text, operation.executed)
      case 'release': {
        const released = requireSession(operation)
        released.release()
        // A release acknowledgement is an idle barrier, not just an abort notification.
        await Promise.all(
          [...pending.values()]
            .filter(
              ({ operation: item }) =>
                (item.kind === 'execute' || item.kind === 'verify') &&
                item.sessionId === released.id &&
                item.runId === released.runId,
            )
            .map((item) => item.done),
        )
        if (session === released) session = null
        return null
      }
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
    let settle!: () => void
    const done = new Promise<void>((resolve) => {
      settle = resolve
    })
    pending.set(request.id, { controller, operation: request.operation, done, settle })
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
        settle()
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
      }
    }
  })
  const offChanges = adapter.subscribeChanges(() => {
    if (!disposed) port.event({ hostId, kind: 'changed' })
  })
  port.event({ hostId, kind: 'ready' })
  return () => {
    if (disposed) return
    disposed = true
    port.event({ hostId, kind: 'disposed' })
    offChanges()
    offRequest()
    offCancel()
    for (const { controller } of pending.values()) controller.abort()
    pending.clear()
    session?.release()
    session = null
    adapter.dispose()
  }
}
