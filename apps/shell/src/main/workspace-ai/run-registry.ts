import type {
  WorkspaceAiOwner,
  WorkspaceAiBoundRun,
  WorkspaceAiRejection,
  WorkspaceAiResult,
  WorkspaceAiRunRequest,
  WorkspaceAiRunState,
  WorkspaceAiStoppedRun,
  WorkspaceAiToolReply,
  WorkspaceAiToolRequest,
  WorkspaceAiToolTicket,
  WorkspaceDocumentEndpoint,
  WorkspaceDocumentIdentity,
} from '../../shared/workspace-ai-api'

interface Run {
  readonly request: WorkspaceAiBoundRun
  state: WorkspaceAiRunState
  revision: string
  deadline: number
  readonly calls: Set<string>
  pending: WorkspaceAiToolTicket | null
}

const fail = (reason: WorkspaceAiRejection): WorkspaceAiResult<never> => ({ ok: false, reason })
const ok = <T>(value: T): WorkspaceAiResult<T> => ({ ok: true, value })
const textId = (value: string): boolean => typeof value === 'string' && value.length > 0
const integer = (value: number): boolean => Number.isSafeInteger(value) && value > 0
const documentKinds = new Set([
  'docs',
  'sheets',
  'slides',
  'pdf',
  'markdown',
  'html',
  'mindmap',
  'hwp',
  'epub',
  'diagram',
])

function sameOwner(a: WorkspaceAiOwner, b: WorkspaceAiOwner): boolean {
  return (
    a.windowId === b.windowId &&
    a.panelContentsId === b.panelContentsId &&
    a.panelGeneration === b.panelGeneration
  )
}

export function sameWorkspaceDocument(
  a: WorkspaceDocumentIdentity,
  b: WorkspaceDocumentIdentity,
): boolean {
  return (
    a.windowId === b.windowId &&
    a.tabId === b.tabId &&
    a.documentId === b.documentId &&
    a.kind === b.kind &&
    a.generation === b.generation
  )
}

/**
 * Trusted main-process state; the execution controller owns editor RPC. One active run
 * per window; FIFO keeps each conversation's queued snapshots in order too.
 * The caller starts a queued run explicitly after applying the existing send gate.
 */
export class WorkspaceAiRunRegistry {
  private readonly owner: WorkspaceAiOwner
  private readonly documents = new Map<string, WorkspaceDocumentEndpoint>()
  private readonly runs = new Map<string, Run>()
  private readonly queue: string[] = []
  private active: string | null = null
  private nextBindingId = 1
  private disposed = false

  constructor(
    owner: WorkspaceAiOwner,
    private readonly now: () => number = Date.now,
    private readonly timeoutMs = 120_000,
    private readonly queueLimit = 32,
    private readonly onStopped: (run: WorkspaceAiStoppedRun) => void = () => {},
    private readonly terminalLimit = 128,
  ) {
    if (
      !integer(owner.windowId) ||
      !integer(owner.panelContentsId) ||
      !integer(owner.panelGeneration) ||
      !Number.isFinite(timeoutMs) ||
      timeoutMs <= 0 ||
      !integer(queueLimit) ||
      !integer(terminalLimit)
    ) {
      throw new Error('invalid workspace owner or run limits')
    }
    this.owner = Object.freeze({ ...owner })
  }

  /** Main calls this only after resolving a live view in the registry's own window. */
  registerDocument(endpoint: WorkspaceDocumentEndpoint): void {
    const document = endpoint.document
    if (
      this.disposed ||
      document.windowId !== this.owner.windowId ||
      !textId(document.tabId) ||
      !textId(document.documentId) ||
      !documentKinds.has(document.kind) ||
      !integer(document.generation) ||
      !integer(endpoint.contentsId) ||
      endpoint.contentsId === this.owner.panelContentsId ||
      !textId(endpoint.revision) ||
      typeof endpoint.readOnly !== 'boolean' ||
      endpoint.tools.some(
        (tool) => !textId(tool.name) || (tool.access !== 'read' && tool.access !== 'write'),
      ) ||
      new Set(endpoint.tools.map((tool) => tool.name)).size !== endpoint.tools.length
    ) {
      throw new Error('invalid workspace document endpoint')
    }
    const previous = this.documents.get(document.tabId)
    if (previous) {
      const replaced =
        !sameWorkspaceDocument(previous.document, document) ||
        previous.contentsId !== endpoint.contentsId ||
        previous.readOnly !== endpoint.readOnly ||
        previous.tools.length !== endpoint.tools.length ||
        previous.tools.some(
          (tool, index) =>
            tool.name !== endpoint.tools[index]?.name ||
            tool.access !== endpoint.tools[index]?.access,
        )
      const changed = previous.revision !== endpoint.revision
      if (replaced || changed) {
        this.invalidateDocument(
          previous.document,
          replaced ? 'editor-replaced' : 'document-changed',
        )
      }
    }
    this.documents.set(document.tabId, {
      ...endpoint,
      document: Object.freeze({ ...document }),
      tools: endpoint.tools.map((tool) => Object.freeze({ ...tool })),
    })
  }

  /** A manual edit must use this path; accepted tool replies advance their own revision. */
  documentChanged(document: WorkspaceDocumentIdentity, revision: string): void {
    const endpoint = this.endpoint(document)
    if (!endpoint || !textId(revision) || revision === endpoint.revision) return
    this.documents.set(document.tabId, { ...endpoint, revision })
    this.invalidateDocument(document, 'document-changed')
  }

  closeDocument(document: WorkspaceDocumentIdentity): void {
    if (!this.endpoint(document)) return
    this.documents.delete(document.tabId)
    this.invalidateDocument(document, 'document-closed')
  }

  enqueue(request: WorkspaceAiRunRequest): WorkspaceAiResult<WorkspaceAiBoundRun> {
    if (!sameOwner(request.owner, this.owner)) return fail('wrong-owner')
    if (this.disposed) return fail('panel-closed')
    if (!textId(request.runId) || !textId(request.conversationId)) return fail('invalid-request')
    if (this.runs.has(request.runId)) return fail('duplicate-run')
    if (this.queue.length >= this.queueLimit) return fail('queue-full')
    const documents = [request.target, ...request.references]
    if (
      documents.some(
        (binding) =>
          !textId(binding.revision) ||
          (binding.selection !== null &&
            (!textId(binding.selection.id) || binding.selection.revision !== binding.revision)),
      ) ||
      request.references.some((reference) => !textId(reference.snapshotId)) ||
      (request.project !== null &&
        (!textId(request.project.id) || !textId(request.project.contextVersion))) ||
      new Set(documents.map(({ document }) => document.tabId)).size !== documents.length
    ) {
      return fail('invalid-request')
    }
    for (const binding of documents) {
      const endpoint = this.endpoint(binding.document)
      if (!endpoint) return fail('missing-document')
      if (endpoint.revision !== binding.revision) return fail('document-changed')
    }
    // Copy before enqueue: a composer changing target/project/selection cannot retarget its run.
    if (!Number.isSafeInteger(this.nextBindingId)) return fail('invalid-request')
    const snapshot = { ...structuredClone(request), bindingId: this.nextBindingId++ }
    const run: Run = {
      request: snapshot,
      state: { status: 'queued' },
      revision: snapshot.target.revision,
      deadline: 0,
      calls: new Set(),
      pending: null,
    }
    this.runs.set(snapshot.runId, run)
    this.queue.push(snapshot.runId)
    return ok(structuredClone(snapshot))
  }

  /** No model invocation occurs here. A selected conversation is deliberately not an argument. */
  startNext(): WorkspaceAiBoundRun | null {
    this.expire()
    if (this.disposed || this.active !== null) return null
    while (this.queue.length > 0) {
      const id = this.queue.shift()!
      const run = this.runs.get(id)!
      if (run.state.status !== 'queued') continue
      const problem = this.validateDocuments(run)
      if (problem) {
        this.stop(run, problem)
        continue
      }
      run.state = { status: 'running' }
      run.deadline = this.now() + this.timeoutMs
      this.active = id
      return structuredClone(run.request)
    }
    return null
  }

  claimTool(request: WorkspaceAiToolRequest): WorkspaceAiResult<WorkspaceAiToolTicket> {
    const found = this.running(
      request.owner,
      request.conversationId,
      request.runId,
      request.bindingId,
    )
    if (!found.ok) return found
    const run = found.value
    const problem = this.validateDocuments(run)
    if (problem) {
      this.stop(run, problem)
      return fail(problem)
    }
    if (!textId(request.callId)) return fail('invalid-request')
    if (run.calls.has(request.callId)) return fail('duplicate-call')
    if (run.pending) return fail('call-in-flight')
    const target = sameWorkspaceDocument(request.document, run.request.target.document)
    const reference = run.request.references.find((ref) =>
      sameWorkspaceDocument(request.document, ref.document),
    )
    if (!target && !reference) return fail('unbound-document')
    const endpoint = this.endpoint(request.document)!
    const tool = endpoint.tools.find((item) => item.name === request.tool)
    if (!tool) return fail('unknown-tool')
    if (tool.access === 'write' && (!target || endpoint.readOnly)) return fail('read-only')
    const binding = target ? run.request.target : reference!
    const ticket: WorkspaceAiToolTicket = {
      ...structuredClone(request),
      contentsId: endpoint.contentsId,
      access: tool.access,
      revision: target ? run.revision : binding.revision,
      selection: structuredClone(binding.selection),
      snapshotId: reference?.snapshotId ?? null,
    }
    run.calls.add(request.callId)
    run.pending = ticket
    return ok(structuredClone(ticket))
  }

  /** senderContentsId comes from the IPC event, not from its reply payload. */
  acceptReply(senderContentsId: number, reply: WorkspaceAiToolReply): WorkspaceAiResult<void> {
    this.expire()
    const run = this.runs.get(reply.runId)
    if (!run || run.state.status !== 'running') return fail('run-finished')
    const ticket = run.pending
    if (
      !ticket ||
      senderContentsId !== ticket.contentsId ||
      !sameOwner(reply.owner, ticket.owner) ||
      reply.bindingId !== ticket.bindingId ||
      reply.conversationId !== ticket.conversationId ||
      reply.callId !== ticket.callId ||
      reply.tool !== ticket.tool ||
      !sameWorkspaceDocument(reply.document, ticket.document)
    ) {
      return fail('wrong-response')
    }
    const problem = this.validateDocuments(run)
    if (problem) {
      this.stop(run, problem)
      return fail(problem)
    }
    if (
      !textId(reply.revision) ||
      typeof reply.mutated !== 'boolean' ||
      (ticket.access === 'read' && (reply.mutated || reply.revision !== ticket.revision)) ||
      (!reply.mutated && reply.revision !== ticket.revision) ||
      (reply.mutated && reply.revision === ticket.revision)
    ) {
      this.stop(run, 'document-changed')
      return fail('document-changed')
    }
    if (reply.mutated) {
      run.revision = reply.revision
      const endpoint = this.endpoint(reply.document)!
      this.documents.set(reply.document.tabId, { ...endpoint, revision: reply.revision })
      // Queued snapshots of this document must be reviewed again after an earlier run edits it.
      this.invalidateDocument(reply.document, 'document-changed', reply.runId)
    }
    run.pending = null
    return ok(undefined)
  }

  complete(
    owner: WorkspaceAiOwner,
    conversationId: string,
    runId: string,
    bindingId: number,
  ): WorkspaceAiResult<void> {
    const found = this.running(owner, conversationId, runId, bindingId)
    if (!found.ok) return found
    const run = found.value
    if (run.pending) return fail('call-in-flight')
    const problem = this.validateDocuments(run)
    if (problem) {
      this.stop(run, problem)
      return fail(problem)
    }
    run.state = { status: 'completed' }
    this.active = null
    this.trimTerminalRuns()
    return ok(undefined)
  }

  cancel(
    owner: WorkspaceAiOwner,
    conversationId: string,
    runId: string,
    bindingId: number,
  ): WorkspaceAiResult<void> {
    return this.stopRun(owner, conversationId, runId, bindingId, 'cancelled')
  }

  /** Trusted host failures stop exactly the bound run, including an in-flight tool. */
  stopRun(
    owner: WorkspaceAiOwner,
    conversationId: string,
    runId: string,
    bindingId: number,
    reason: WorkspaceAiRejection,
  ): WorkspaceAiResult<void> {
    if (!sameOwner(owner, this.owner)) return fail('wrong-owner')
    const run = this.runs.get(runId)
    if (
      !run ||
      run.request.conversationId !== conversationId ||
      run.request.bindingId !== bindingId
    ) {
      return fail('run-not-active')
    }
    if (run.state.status !== 'running' && run.state.status !== 'queued') return fail('run-finished')
    this.stop(run, reason)
    return ok(undefined)
  }

  state(runId: string): WorkspaceAiRunState | null {
    this.expire()
    const state = this.runs.get(runId)?.state
    return state ? { ...state } : null
  }

  /** onStopped fires even when state/claimTool discovers expiry; the host cancels live work there. */
  expire(): string[] {
    const run = this.active ? this.runs.get(this.active) : null
    if (!run || this.now() < run.deadline) return []
    this.stop(run, 'timeout')
    return [run.request.runId]
  }

  /** Panel reload/close creates a new registry; old tickets can never enter the new generation. */
  dispose(): void {
    this.disposed = true
    for (const run of this.runs.values()) {
      if (run.state.status === 'queued' || run.state.status === 'running') {
        this.stop(run, 'panel-closed')
      }
    }
    this.documents.clear()
  }

  private endpoint(document: WorkspaceDocumentIdentity): WorkspaceDocumentEndpoint | undefined {
    const endpoint = this.documents.get(document.tabId)
    return endpoint && sameWorkspaceDocument(endpoint.document, document) ? endpoint : undefined
  }

  private running(
    owner: WorkspaceAiOwner,
    conversationId: string,
    runId: string,
    bindingId: number,
  ): WorkspaceAiResult<Run> {
    if (!sameOwner(owner, this.owner)) return fail('wrong-owner')
    this.expire()
    const run = this.runs.get(runId)
    if (
      !run ||
      run.request.conversationId !== conversationId ||
      run.request.bindingId !== bindingId
    ) {
      return fail('run-not-active')
    }
    if (run.state.status === 'stopped') return fail(run.state.reason)
    if (run.state.status !== 'running' || this.active !== runId) return fail('run-not-active')
    return ok(run)
  }

  private validateDocuments(run: Run): WorkspaceAiRejection | null {
    for (const binding of [run.request.target, ...run.request.references]) {
      const endpoint = this.endpoint(binding.document)
      if (!endpoint) return 'missing-document'
      const revision = binding === run.request.target ? run.revision : binding.revision
      if (endpoint.revision !== revision) return 'document-changed'
    }
    return null
  }

  private stop(run: Run, reason: WorkspaceAiRejection): void {
    run.state = { status: 'stopped', reason }
    run.pending = null
    if (this.active === run.request.runId) this.active = null
    const index = this.queue.indexOf(run.request.runId)
    if (index !== -1) this.queue.splice(index, 1)
    this.trimTerminalRuns()
    this.onStopped({
      conversationId: run.request.conversationId,
      runId: run.request.runId,
      bindingId: run.request.bindingId,
      reason,
    })
  }

  private trimTerminalRuns(): void {
    const terminal = [...this.runs.values()].filter(
      (run) => run.state.status === 'completed' || run.state.status === 'stopped',
    )
    for (const run of terminal.slice(0, -this.terminalLimit)) {
      this.runs.delete(run.request.runId)
    }
  }

  private invalidateDocument(
    document: WorkspaceDocumentIdentity,
    reason: WorkspaceAiRejection,
    exceptRunId?: string,
  ): void {
    for (const run of this.runs.values()) {
      if (
        run.request.runId !== exceptRunId &&
        (run.state.status === 'queued' || run.state.status === 'running') &&
        [run.request.target, ...run.request.references].some((binding) =>
          sameWorkspaceDocument(binding.document, document),
        )
      ) {
        this.stop(run, reason)
      }
    }
  }
}
