import type { WorkspaceReadRequest } from '@genoffice/agent-core'
import type { WorkspaceConversationOrigin } from '../../../shared/workspace-ai-api'
import type { WorkspaceHistoryMessage } from '../../../main/workspace-ai/conversation-history'
import type {
  WorkspacePanelApi,
  WorkspacePanelEvent,
  WorkspacePanelReply,
  WorkspacePanelRequest,
  WorkspacePanelSession,
  WorkspacePanelValues,
} from '../../../shared/workspace-panel-api'
import type { WorkspacePanelLayout } from '../../../shared/workspace-layout'
import type { WorkspaceStartedRun } from '../../../main/workspace-ai/execution-controller'
import {
  WorkspaceRunHost,
  type WorkspaceExecutionBridge,
  type WorkspaceRunHostDeps,
} from './run-host'
import { WorkspacePanelController } from './panel-controller'

function unwrap<T>(reply: WorkspacePanelReply<T>): T {
  if (!reply.ok) throw new Error(`Workspace panel rejected: ${reply.reason}`)
  return reply.value
}

/** Restores process-local abort signals; no DOM or provider dependency. */
export class WorkspacePanelClient implements WorkspaceExecutionBridge {
  private readonly signals = new Map<number, AbortController>()
  private readonly stopped = new Set<number>()
  private readonly listeners = new Set<(event: WorkspacePanelEvent) => void>()
  private closed = false
  private off: () => void = () => {}
  private constructor(
    private readonly api: WorkspacePanelApi,
    readonly session: WorkspacePanelSession,
  ) {}

  static async connect(api: WorkspacePanelApi): Promise<WorkspacePanelClient> {
    const client = new WorkspacePanelClient(api, unwrap(await api.connect()))
    client.off = api.onEvent((event) => client.receive(event))
    return client
  }
  private receive(event: WorkspacePanelEvent): void {
    if (this.closed || event.sessionId !== this.session.id) return
    if (event.kind === 'stopped') {
      this.stopped.add(event.run.bindingId)
      if (this.stopped.size > 128) this.stopped.delete(this.stopped.values().next().value!)
      this.signals.get(event.run.bindingId)?.abort()
      this.signals.delete(event.run.bindingId)
    }
    if (event.kind === 'closed') this.invalidate()
    for (const listener of [...this.listeners]) listener(event)
  }
  private invalidate(): void {
    this.closed = true
    for (const signal of this.signals.values()) signal.abort()
    this.signals.clear()
    this.stopped.clear()
    this.off()
  }
  subscribe(listener: (event: WorkspacePanelEvent) => void): () => void {
    this.listeners.add(listener)
    return () => {
      this.listeners.delete(listener)
    }
  }
  private async call<R extends WorkspacePanelRequest>(
    request: R,
  ): Promise<WorkspacePanelValues[R['kind']]> {
    if (this.closed) throw new Error('Workspace panel closed')
    const reply = await this.api.request(this.session.id, request)
    if (this.closed) throw new Error('Workspace panel closed')
    return unwrap(reply)
  }
  catalog = () => this.call({ kind: 'catalog' })
  layout = (panel: WorkspacePanelLayout | null) => this.call({ kind: 'layout', panel })
  capture = (tabId: string) => this.call({ kind: 'capture', tabId })
  reference = (captureId: string, scope: WorkspaceReadRequest['scope']) =>
    this.call({ kind: 'reference', captureId, scope })
  create = () => this.call({ kind: 'create' })
  load = (origin: WorkspaceConversationOrigin) => this.call({ kind: 'load', origin })
  append = (origin: WorkspaceConversationOrigin, message: WorkspaceHistoryMessage) =>
    this.call({ kind: 'append', origin, message })
  enqueue: WorkspaceExecutionBridge['enqueue'] = (request) =>
    this.call({ kind: 'enqueue', request })
  async startNext(): Promise<WorkspaceStartedRun | null> {
    const run = await this.call({ kind: 'startNext' })
    if (!run) return null
    const abort = new AbortController()
    const id = run.binding.bindingId
    if (run.aborted || this.stopped.has(id)) abort.abort()
    else this.signals.set(id, abort)
    return {
      binding: run.binding,
      target: run.target,
      references: run.references,
      signal: abort.signal,
    }
  }
  execute: WorkspaceExecutionBridge['execute'] = (...args) => this.call({ kind: 'execute', args })
  verifyResponse: WorkspaceExecutionBridge['verifyResponse'] = (...args) =>
    this.call({ kind: 'verifyResponse', args })
  complete: WorkspaceExecutionBridge['complete'] = async (...args) => {
    await this.call({ kind: 'complete', args })
    this.signals.delete(args[0].bindingId)
  }
  cancel: WorkspaceExecutionBridge['cancel'] = async (...args) => {
    // Main still validates the complete owner/run handle before touching any editor.
    await this.call({ kind: 'cancel', args })
  }
  state: WorkspaceExecutionBridge['state'] = (runId) => this.call({ kind: 'state', runId })
  rollbackForRun: WorkspaceExecutionBridge['rollbackForRun'] = (...args) =>
    this.call({ kind: 'rollbackForRun', args })
  rollback: WorkspaceExecutionBridge['rollback'] = (...args) =>
    this.call({ kind: 'rollback', args })
  async dispose(): Promise<void> {
    if (this.closed) return
    this.invalidate()
    this.listeners.clear()
    unwrap(await this.api.request(this.session.id, { kind: 'disconnect' }))
  }
}

/** The caller must supply the existing gate, guarded preparation and transport explicitly. */
export async function connectWorkspacePanel(
  api: WorkspacePanelApi,
  deps: Omit<WorkspaceRunHostDeps, 'execution' | 'load' | 'append'>,
) {
  const client = await WorkspacePanelClient.connect(api)
  const host = new WorkspaceRunHost({
    ...deps,
    execution: client,
    load: client.load,
    append: client.append,
  })
  const controller = new WorkspacePanelController({
    host,
    capture: client.capture,
    reference: client.reference,
    create: client.create,
  })
  let closed = false
  let epoch = 0
  const refresh = async () => {
    const current = ++epoch
    const catalog = await client.catalog()
    if (!closed && epoch === current) controller.catalog(catalog.documents, catalog.sources)
  }
  const off = client.subscribe((event) => {
    if (event.kind === 'history-changed') host.historyChanged(event.origin)
    if (event.kind === 'history-rebound') host.rebindHistory(event.previous, event.origin)
    if (event.kind === 'catalog-changed') void refresh().catch(() => {})
    if (event.kind === 'closed') {
      closed = true
      controller.dispose()
      host.dispose()
      off()
    }
  })
  try {
    await refresh()
  } catch (error) {
    off()
    controller.dispose()
    host.dispose()
    await client.dispose().catch(() => {})
    throw error
  }
  return {
    client,
    host,
    controller,
    refresh,
    async dispose() {
      closed = true
      off()
      controller.dispose()
      host.dispose()
      await client.dispose()
    },
  }
}
