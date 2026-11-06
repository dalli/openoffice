import { randomUUID } from 'node:crypto'
import type {
  AgentToolCall,
  ExecutedToolCall,
  WorkspaceAdapterState,
  WorkspaceReadRequest,
  WorkspaceReadSnapshot,
  WorkspaceTemplateMode,
  WorkspaceWriterUpdate,
  WorkspaceTemplateSnapshot,
} from '@genoffice/agent-core'
import type {
  WorkspaceEditorDescription,
  WorkspaceEditorOperation,
  WorkspaceEditorSession,
  WorkspaceEditorValues,
  WorkspaceEditQueueCommand,
  WorkspaceEditQueueSnapshot,
} from '@genoffice/agent-core/workspace-editor-protocol'
import type { DocumentTabKind } from '../../shared/tabs-api'
import type {
  WorkspaceAiBoundRun,
  WorkspaceAiOwner,
  WorkspaceAiRejection,
  WorkspaceAiResult,
  WorkspaceAiRunState,
  WorkspaceAiStoppedRun,
  WorkspaceAiToolTicket,
  WorkspaceBoundDocument,
  WorkspaceDocumentIdentity,
  WorkspaceProjectBinding,
  WorkspaceReference,
} from '../../shared/workspace-ai-api'
import { sameWorkspaceDocument, WorkspaceAiRunRegistry } from './run-registry'

/** Only the owning main process can resolve these clients from its live tab views. */
export interface WorkspaceEditorClient {
  readonly contentsId: number
  readonly available: boolean
  request<O extends WorkspaceEditorOperation>(
    operation: O,
    signal?: AbortSignal,
  ): Promise<WorkspaceEditorValues[O['kind']]>
  onChanged(handler: () => void): () => void
  onQueueChanged?(handler: () => void): () => void
  onClosed(handler: () => void): () => void
}

export interface WorkspaceExecutionDeps {
  readonly owner: WorkspaceAiOwner
  resolve(tabId: string): { kind: DocumentTabKind; client: WorkspaceEditorClient } | null
  onStopped?(run: WorkspaceAiStoppedRun): void
  onDisposed?(): void
  onQueueChanged?(tabId: string): void
  now?(): number
  readonly timeoutMs?: number
}

export interface WorkspaceCapture {
  readonly id: string
  readonly binding: WorkspaceBoundDocument
}

export interface WorkspaceReferenceCapture {
  readonly id: string
  readonly binding: WorkspaceReference
  readonly snapshot: WorkspaceReadSnapshot
}

export interface WorkspaceRunHandle {
  readonly owner: WorkspaceAiOwner
  readonly conversationId: string
  readonly runId: string
  readonly bindingId: number
}

export interface WorkspaceStartedRun {
  readonly binding: WorkspaceAiBoundRun
  readonly target: WorkspaceEditorSession
  readonly references: readonly WorkspaceReferenceCapture[]
  /** The model host must bind its existing transport cancellation to this signal. */
  readonly signal: AbortSignal
}

/** Main-issued capability; the snapshot and the editor's own token never leave their owners. */
export interface WorkspaceRollbackHandle extends WorkspaceRunHandle {
  readonly rollbackId: string
}

interface RollbackEntry {
  handle: WorkspaceRollbackHandle
  resource: Resource
  document: WorkspaceDocumentIdentity
  state: WorkspaceAdapterState
  editorToken: string
}

interface Resource {
  readonly tabId: string
  readonly kind: DocumentTabKind
  readonly client: WorkspaceEditorClient
  identity: WorkspaceDocumentIdentity | null
  description: WorkspaceEditorDescription | null
  serial: number
  off(): void
}
interface Captured {
  readonly capture: WorkspaceCapture
  readonly resource: Resource
  readonly state: WorkspaceAdapterState
  readonly serial: number
}
interface ReferenceEntry extends Captured {
  readonly reference: WorkspaceReferenceCapture
}
interface QueuedRun {
  readonly editorRunId: string
  readonly binding: WorkspaceAiBoundRun
  readonly target: Captured
  readonly references: readonly ReferenceEntry[]
  readonly abort: AbortController
  expected: WorkspaceAdapterState
  session: WorkspaceEditorSession | null
  pending: WorkspaceAiToolTicket | null
  writerReady?: Promise<void>
  busy: boolean
  timer: ReturnType<typeof setTimeout> | null
}

export class WorkspaceExecutionError extends Error {
  constructor(readonly kind: WorkspaceAiRejection) {
    super(`Workspace execution rejected: ${kind}`)
    this.name = 'WorkspaceExecutionError'
  }
}
function fail(kind: WorkspaceAiRejection): never {
  throw new WorkspaceExecutionError(kind)
}
const unwrap = <T>(result: WorkspaceAiResult<T>): T =>
  result.ok ? result.value : fail(result.reason)
const sameVersion = (a: WorkspaceAdapterState, b: WorkspaceAdapterState) =>
  a.instanceId === b.instanceId && a.revision === b.revision && a.readOnly === b.readOnly
const bindingOf = (
  document: WorkspaceDocumentIdentity,
  state: WorkspaceAdapterState,
): WorkspaceBoundDocument => ({
  document: { ...document },
  revision: state.revision,
  selection: state.selectionId ? { id: state.selectionId, revision: state.revision } : null,
})

/**
 * Main-only document coordinator. It never selects a provider, invokes a model,
 * reads a file or persists a conversation. A future panel still uses the existing
 * send gate and project policy before it may start its model transport.
 */
export class WorkspaceExecutionController {
  private readonly owner: WorkspaceAiOwner
  private readonly registry: WorkspaceAiRunRegistry
  private readonly resources = new Map<string, Resource>()
  private readonly captures = new Map<string, Captured>()
  private readonly snapshots = new Map<string, ReferenceEntry>()
  private readonly runs = new Map<number, QueuedRun>()
  private readonly rollbacks = new Map<number, RollbackEntry>()
  private readonly lifetime = new AbortController()
  private readonly timeoutMs: number
  private generation = 0
  private disposed = false
  private starting = false
  private rollingBack = false
  private savingCopy: { handle: WorkspaceRunHandle; abort: AbortController } | null = null
  private active: QueuedRun | null = null
  private draining: Promise<void> = Promise.resolve()

  constructor(private readonly deps: WorkspaceExecutionDeps) {
    this.owner = Object.freeze({ ...deps.owner })
    this.timeoutMs = deps.timeoutMs ?? 120_000
    this.registry = new WorkspaceAiRunRegistry(
      this.owner,
      deps.now ?? Date.now,
      this.timeoutMs,
      32,
      (run) => this.stopped(run),
    )
  }

  async capture(tabId: string): Promise<WorkspaceCapture> {
    return this.captureTarget(tabId)
  }

  /** Queued passages have their own anchors; the last visible selection must not scope the batch. */
  async captureEditBatch(
    tabId: string,
    expected: { documentId: string; revision: string },
  ): Promise<WorkspaceCapture> {
    const capture = await this.capture(tabId)
    const entry = this.captures.get(capture.id) ?? fail('invalid-request')
    if (
      entry.resource.kind !== 'docs' ||
      entry.state.instanceId !== expected.documentId ||
      entry.state.revision !== expected.revision ||
      entry.state.readOnly
    )
      fail('document-changed')
    const whole: Captured = {
      ...entry,
      state: { ...entry.state, selectionId: null },
      capture: {
        id: randomUUID(),
        binding: { ...capture.binding, selection: null },
      },
    }
    this.captures.delete(capture.id)
    this.captures.set(whole.capture.id, whole)
    return structuredClone(whole.capture)
  }

  /** An explicit card action reads the original target's current edited copy, never the focused tab or disk. */
  async template(
    document: WorkspaceDocumentIdentity,
    mode: WorkspaceTemplateMode,
  ): Promise<WorkspaceTemplateSnapshot | null> {
    if (document.windowId !== this.owner.windowId) fail('wrong-owner')
    if (this.active) fail('call-in-flight')
    const capture = await this.capture(document.tabId)
    if (!sameWorkspaceDocument(capture.binding.document, document)) fail('editor-replaced')
    const entry = this.captures.get(capture.id) ?? fail('invalid-request')
    const serial = entry.resource.serial
    const snapshot = await this.request(
      entry.resource,
      { kind: 'template', state: entry.state, mode },
      this.lifetime.signal,
    )
    this.live()
    if (this.active || entry.resource.serial !== serial) fail('document-changed')
    const state = await this.request(entry.resource, { kind: 'capture' }, this.lifetime.signal)
    if (entry.resource.serial !== serial || !sameVersion(state, entry.state))
      fail('document-changed')
    return snapshot
  }

  /** UI-only queue access stays bound to the selected document's live view. */
  async editQueue(
    document: WorkspaceDocumentIdentity,
    command: WorkspaceEditQueueCommand,
  ): Promise<WorkspaceEditQueueSnapshot | null> {
    this.live()
    if (document.windowId !== this.owner.windowId) fail('wrong-owner')
    if (document.kind !== 'docs') fail('invalid-request')
    if (this.active || this.starting || this.rollingBack || this.savingCopy) fail('call-in-flight')
    const resource = this.resource(document.tabId)
    if (!resource.identity || !sameWorkspaceDocument(resource.identity, document))
      fail('editor-replaced')
    const state = await this.request(resource, { kind: 'capture' }, this.lifetime.signal)
    this.live()
    if (state.instanceId !== document.documentId || this.resource(document.tabId) !== resource)
      fail('editor-replaced')
    if (this.active || this.starting || this.rollingBack || this.savingCopy) fail('call-in-flight')
    // Reading a card must not allocate captures and evict another conversation's target.
    return this.request(resource, { kind: 'edit-queue', state, command }, this.lifetime.signal)
  }

  /** Renew only this document and its retained range. References and queued bindings stay fixed. */
  async refresh(captureId: string): Promise<WorkspaceCapture> {
    this.live()
    const previous = this.captures.get(captureId) ?? fail('invalid-request')
    if (
      this.resource(previous.resource.tabId) !== previous.resource ||
      !previous.resource.identity ||
      !sameWorkspaceDocument(previous.capture.binding.document, previous.resource.identity)
    )
      fail('editor-replaced')
    return this.captureTarget(previous.resource.tabId, previous)
  }

  private async captureTarget(tabId: string, previous?: Captured): Promise<WorkspaceCapture> {
    this.live()
    const resource = this.resource(tabId)
    if (this.active?.pending?.access === 'write' && this.active.target.resource === resource)
      fail('call-in-flight')
    const serial = resource.serial
    let description = await this.request(resource, { kind: 'describe' }, this.lifetime.signal)
    if (previous) {
      if (description.state.instanceId !== previous.state.instanceId) fail('editor-replaced')
      const state = await this.request(
        resource,
        { kind: 'refresh', state: previous.state },
        this.lifetime.signal,
      )
      if (
        !sameVersion(description.state, state) ||
        state.selectionId !== previous.state.selectionId
      )
        fail('document-changed')
      description = { ...description, state }
    }
    this.live()
    if (serial !== resource.serial) fail('document-changed')
    if (!description.observesChanges) fail('unsupported-read')
    const identity = resource.identity
    if (!identity || identity.documentId !== description.state.instanceId) {
      resource.identity = {
        windowId: this.owner.windowId,
        tabId,
        kind: resource.kind,
        documentId: description.state.instanceId,
        generation: ++this.generation,
      }
    }
    resource.description = structuredClone(description)
    this.registry.registerDocument({
      document: resource.identity!,
      contentsId: resource.client.contentsId,
      revision: description.state.revision,
      readOnly: description.state.readOnly,
      tools: Object.entries(description.toolAccess).map(([name, access]) => ({ name, access })),
    })
    const captured: Captured = {
      capture: { id: randomUUID(), binding: bindingOf(resource.identity!, description.state) },
      resource,
      state: { ...description.state },
      serial,
    }
    this.remember(this.captures, captured.capture.id, captured, 128)
    return structuredClone(captured.capture)
  }

  /** Explicit current-content acquisition; never reads a document merely because it is visible. */
  async readReference(
    captureId: string,
    read: WorkspaceReadRequest,
  ): Promise<WorkspaceReferenceCapture> {
    this.live()
    const captured = this.captures.get(captureId) ?? fail('invalid-request')
    this.validCapture(captured)
    const snapshot = await this.request(
      captured.resource,
      { kind: 'read', state: captured.state, read: structuredClone(read) },
      this.lifetime.signal,
    )
    this.validCapture(captured)
    const id = randomUUID()
    const reference: WorkspaceReferenceCapture = {
      id,
      binding: { ...structuredClone(captured.capture.binding), snapshotId: id },
      snapshot: structuredClone(snapshot),
    }
    this.remember(this.snapshots, id, { ...captured, reference }, 32)
    return structuredClone(reference)
  }

  enqueue(request: {
    conversationId: string
    runId: string
    targetId: string
    referenceIds: readonly string[]
    project: WorkspaceProjectBinding | null
  }): WorkspaceAiBoundRun {
    this.live()
    const target = this.captures.get(request.targetId) ?? fail('invalid-request')
    const references = request.referenceIds.map(
      (id) => this.snapshots.get(id) ?? fail('invalid-request'),
    )
    if (
      new Set(request.referenceIds).size !== references.length ||
      references.length > 20 ||
      references.reduce((sum, item) => sum + item.reference.snapshot.text.length, 0) > 200_000
    )
      fail('invalid-request')
    for (const captured of [target, ...references]) this.validCapture(captured)
    const binding = unwrap(
      this.registry.enqueue({
        owner: this.owner,
        conversationId: request.conversationId,
        runId: request.runId,
        target: target.capture.binding,
        references: references.map((item) => item.reference.binding),
        project: request.project,
      }),
    )
    this.runs.set(binding.bindingId, {
      editorRunId: randomUUID(),
      binding,
      target,
      references,
      abort: new AbortController(),
      expected: { ...target.state },
      session: null,
      pending: null,
      busy: false,
      timer: null,
    })
    return structuredClone(binding)
  }

  /** FIFO preparation only. The caller must apply the existing send gate before any model call. */
  async startNext(): Promise<WorkspaceStartedRun | null> {
    this.live()
    if (this.starting || this.rollingBack || this.savingCopy) return null
    this.starting = true
    try {
      await this.draining
      this.live()
      const binding = this.registry.startNext()
      if (!binding) return null
      const run = this.runs.get(binding.bindingId)!
      this.active = run
      run.timer = setTimeout(() => this.registry.expire(), this.timeoutMs)
      try {
        await this.checkDocuments(run)
        run.session = await this.request(
          run.target.resource,
          {
            kind: 'prepare',
            request: { runId: run.editorRunId, mode: 'edit', state: run.expected },
          },
          run.abort.signal,
        )
        this.running(binding)
        await this.checkDocuments(run)
        return {
          binding: structuredClone(binding),
          target: structuredClone(run.session),
          references: run.references.map((item) => structuredClone(item.reference)),
          signal: run.abort.signal,
        }
      } catch (error) {
        this.stopForError(run, error)
        throw error
      }
    } finally {
      this.starting = false
    }
  }

  async execute(
    handle: WorkspaceRunHandle,
    document: WorkspaceDocumentIdentity,
    raw: AgentToolCall,
  ): Promise<WorkspaceEditorValues['execute']> {
    const run = this.running(handle)
    if (run.busy) fail('call-in-flight')
    const call = structuredClone(raw)
    const ticket = unwrap(
      this.registry.claimTool({
        owner: this.owner,
        conversationId: run.binding.conversationId,
        runId: run.binding.runId,
        bindingId: run.binding.bindingId,
        document,
        tool: call.name,
        callId: call.id,
      }),
    )
    // Reference snapshots are scoped material, not authority to read arbitrary extra blocks.
    if (!sameWorkspaceDocument(document, run.binding.target.document)) {
      this.registry.stopRun(
        this.owner,
        handle.conversationId,
        handle.runId,
        handle.bindingId,
        'unknown-tool',
      )
      fail('unknown-tool')
    }
    run.busy = true
    let ready!: () => void
    run.writerReady = new Promise<void>((resolve) => {
      ready = resolve
    })
    run.pending = ticket
    try {
      const session = run.session ?? fail('run-not-active')
      if (!session.tools.some((tool) => tool.name === call.name)) fail('unknown-tool')
      await this.checkDocuments(run)
      this.running(handle)
      const executing = this.request(
        run.target.resource,
        { kind: 'execute', sessionId: session.id, runId: session.runId, call },
        run.abort.signal,
      )
      ready()
      const answer = await executing
      this.running(handle)
      // Native editor guards protect the actual mutation; references are checked again before acceptance.
      await this.checkReferences(run)
      if (
        answer.state.instanceId !== run.expected.instanceId ||
        answer.state.selectionId !== run.expected.selectionId
      )
        fail('wrong-response')
      unwrap(
        this.registry.acceptReply(run.target.resource.client.contentsId, {
          ...ticket,
          revision: answer.state.revision,
          mutated: answer.result.mutated === true,
        }),
      )
      run.expected = { ...answer.state }
      if (answer.result.mutated && answer.rollbackId) {
        const previous = this.rollbacks.get(handle.bindingId)
        if (previous && previous.editorToken !== answer.rollbackId) fail('wrong-response')
        this.rollbacks.set(handle.bindingId, {
          handle: previous?.handle ?? { ...structuredClone(handle), rollbackId: randomUUID() },
          resource: run.target.resource,
          document: structuredClone(run.binding.target.document),
          state: { ...answer.state },
          editorToken: answer.rollbackId,
        })
        while (this.rollbacks.size > 16) this.rollbacks.delete(this.rollbacks.keys().next().value!)
      }
      const { rollbackId: _editorToken, ...reply } = answer
      return structuredClone(reply)
    } catch (error) {
      this.stopForError(run, error)
      throw error
    } finally {
      ready()
      run.writerReady = undefined
      run.pending = null
      run.busy = false
    }
  }

  async updateWriter(
    handle: WorkspaceRunHandle,
    callId: string,
    update: WorkspaceWriterUpdate,
  ): Promise<void> {
    const run = this.running(handle)
    if (!run.pending || run.pending.callId !== callId || run.pending.tool !== 'write_document')
      fail('invalid-request')
    await run.writerReady
    this.running(handle)
    if (!run.pending || run.pending.callId !== callId) fail('invalid-request')
    await this.checkReferences(run)
    const session = run.session ?? fail('run-not-active')
    await this.request(
      run.target.resource,
      {
        kind: 'writer-update',
        sessionId: session.id,
        runId: session.runId,
        callId,
        update,
      },
      run.abort.signal,
    )
  }

  async verifyResponse(
    handle: WorkspaceRunHandle,
    text: string,
    executed: readonly ExecutedToolCall[],
  ): Promise<string | null> {
    const run = this.running(handle)
    if (run.busy) fail('call-in-flight')
    run.busy = true
    try {
      await this.checkDocuments(run)
      const session = run.session ?? fail('run-not-active')
      const result = await this.request(
        run.target.resource,
        {
          kind: 'verify',
          sessionId: session.id,
          runId: session.runId,
          text,
          executed: structuredClone(executed),
        },
        run.abort.signal,
      )
      this.running(handle)
      await this.checkDocuments(run)
      return result
    } catch (error) {
      this.stopForError(run, error)
      throw error
    } finally {
      run.busy = false
    }
  }

  async complete(handle: WorkspaceRunHandle): Promise<void> {
    const run = this.running(handle)
    if (run.busy) fail('call-in-flight')
    run.busy = true
    try {
      await this.checkDocuments(run)
      unwrap(
        this.registry.complete(this.owner, handle.conversationId, handle.runId, handle.bindingId),
      )
      this.finish(run)
      await this.draining
    } catch (error) {
      this.stopForError(run, error)
      throw error
    } finally {
      run.busy = false
    }
  }

  /** Only main's acknowledged template capability may request this first project save. */
  async completeProjectCopy(handle: WorkspaceRunHandle, projectId: string): Promise<void> {
    const run = this.running(handle)
    if (
      this.savingCopy ||
      run.binding.project?.id !== projectId ||
      run.target.resource.kind !== 'docs'
    )
      fail('invalid-request')
    const saving = { handle: structuredClone(handle), abort: new AbortController() }
    this.savingCopy = saving
    try {
      await this.complete(handle)
      this.live()
      const saved = await this.request(
        run.target.resource,
        {
          kind: 'save-project-copy',
          runId: run.editorRunId,
          state: run.expected,
          projectId,
        },
        AbortSignal.any([this.lifetime.signal, saving.abort.signal]),
      )
      if (!saved) fail('editor-error')
    } finally {
      if (run.target.resource.client.available) {
        const settling = run.target.resource.client
          .request({ kind: 'settle', runId: run.editorRunId })
          .then(
            () => {},
            () => {},
          )
        this.draining = Promise.all([this.draining, settling]).then(() => {})
        await this.draining
      }
      this.savingCopy = null
    }
  }

  cancel(handle: WorkspaceRunHandle): void {
    this.owned(handle)
    const saving = this.savingCopy
    if (
      saving &&
      saving.handle.bindingId === handle.bindingId &&
      saving.handle.runId === handle.runId &&
      saving.handle.conversationId === handle.conversationId
    ) {
      saving.abort.abort()
      return
    }
    unwrap(this.registry.cancel(this.owner, handle.conversationId, handle.runId, handle.bindingId))
  }

  rollbackForRun(handle: WorkspaceRunHandle): WorkspaceRollbackHandle | null {
    this.owned(handle)
    const saved = this.rollbacks.get(handle.bindingId)
    if (
      !saved ||
      saved.handle.conversationId !== handle.conversationId ||
      saved.handle.runId !== handle.runId
    )
      return null
    return structuredClone(saved.handle)
  }

  async rollback(handle: WorkspaceRollbackHandle): Promise<void> {
    this.live()
    this.owned(handle)
    if (this.rollingBack || this.starting || this.active || this.savingCopy) fail('call-in-flight')
    const saved = this.rollbacks.get(handle.bindingId)
    if (
      !saved ||
      saved.handle.rollbackId !== handle.rollbackId ||
      saved.handle.conversationId !== handle.conversationId ||
      saved.handle.runId !== handle.runId
    )
      fail('wrong-response')
    this.rollingBack = true
    try {
      await this.draining
      this.live()
      if (
        this.resources.get(saved!.resource.tabId) !== saved!.resource ||
        !saved!.resource.identity ||
        !sameWorkspaceDocument(saved!.document, saved!.resource.identity)
      )
        fail('editor-replaced')
      const state = await this.request(saved!.resource, { kind: 'capture' }, this.lifetime.signal)
      if (!sameVersion(saved!.state, state)) fail('document-changed')
      if (state.readOnly) fail('read-only')
      // Consume before dispatch: a failed or cancelled restore is never replayed automatically.
      this.rollbacks.delete(handle.bindingId)
      await this.request(
        saved!.resource,
        { kind: 'rollback', rollbackId: saved!.editorToken, state },
        this.lifetime.signal,
      )
    } finally {
      this.rollingBack = false
    }
  }
  state(runId: string): WorkspaceAiRunState | null {
    return this.registry.state(runId)
  }

  dispose(): void {
    if (this.disposed) return
    this.disposed = true
    this.lifetime.abort()
    this.registry.dispose()
    for (const resource of this.resources.values()) resource.off()
    this.resources.clear()
    this.captures.clear()
    this.snapshots.clear()
    this.rollbacks.clear()
    this.deps.onDisposed?.()
  }

  private live(): void {
    if (this.disposed) fail('panel-closed')
  }
  private running(handle: WorkspaceRunHandle): QueuedRun {
    this.live()
    this.owned(handle)
    const state = this.registry.state(handle.runId)
    const run = this.active
    if (
      !run ||
      run.binding.bindingId !== handle.bindingId ||
      run.binding.runId !== handle.runId ||
      run.binding.conversationId !== handle.conversationId
    )
      fail('run-not-active')
    if (state?.status !== 'running' || !run.session)
      fail(state?.status === 'stopped' ? state.reason : 'run-not-active')
    return run
  }

  private owned(handle: WorkspaceRunHandle): void {
    const owner = handle.owner
    if (
      !owner ||
      owner.windowId !== this.owner.windowId ||
      owner.panelContentsId !== this.owner.panelContentsId ||
      owner.panelGeneration !== this.owner.panelGeneration
    )
      fail('wrong-owner')
  }

  private resource(tabId: string): Resource {
    const resolved = this.deps.resolve(tabId) ?? fail('missing-document')
    if (!resolved.client.available || resolved.client.contentsId === this.owner.panelContentsId)
      fail('missing-document')
    const previous = this.resources.get(tabId)
    if (previous?.client === resolved.client && previous.kind === resolved.kind) return previous
    if (previous) {
      previous.off()
      if (previous.identity) this.registry.closeDocument(previous.identity)
    }
    const resource: Resource = {
      tabId,
      ...resolved,
      identity: null,
      description: null,
      serial: 0,
      off: () => {},
    }
    const offChanges = resolved.client.onChanged(() => {
      resource.serial++
      const writing =
        this.active?.target.resource === resource && this.active?.pending?.access === 'write'
      if (!writing && resource.identity)
        this.registry.documentChanged(resource.identity, `invalidated:${resource.serial}`)
    })
    const offClosed = resolved.client.onClosed(() => {
      resource.serial++
      if (resource.identity) this.registry.closeDocument(resource.identity)
    })
    const offQueue = resolved.client.onQueueChanged?.(() => this.deps.onQueueChanged?.(tabId))
    resource.off = () => {
      offChanges()
      offClosed()
      offQueue?.()
    }
    this.resources.set(tabId, resource)
    return resource
  }

  private validCapture(captured: Captured): void {
    this.live()
    if (
      this.resources.get(captured.resource.tabId) !== captured.resource ||
      !captured.resource.client.available
    )
      fail('missing-document')
    if (
      captured.resource.serial !== captured.serial ||
      !captured.resource.identity ||
      !sameWorkspaceDocument(captured.capture.binding.document, captured.resource.identity)
    )
      fail('document-changed')
  }

  private async checkReferences(run: QueuedRun): Promise<void> {
    for (const reference of run.references) {
      this.validCapture(reference)
      const state = await this.request(reference.resource, { kind: 'capture' }, run.abort.signal)
      this.validCapture(reference)
      if (!sameVersion(reference.state, state)) fail('document-changed')
    }
  }
  private async checkDocuments(run: QueuedRun): Promise<void> {
    const resource = run.target.resource
    if (
      !resource.identity ||
      !sameWorkspaceDocument(resource.identity, run.binding.target.document)
    )
      fail('editor-replaced')
    const state = await this.request(resource, { kind: 'capture' }, run.abort.signal)
    if (!sameVersion(run.expected, state)) fail('document-changed')
    await this.checkReferences(run)
    if (run.abort.signal.aborted) fail('cancelled')
  }

  private async request<O extends WorkspaceEditorOperation>(
    resource: Resource,
    operation: O,
    signal: AbortSignal,
  ): Promise<WorkspaceEditorValues[O['kind']]> {
    if (this.resources.get(resource.tabId) !== resource || !resource.client.available)
      fail('missing-document')
    try {
      return await resource.client.request(operation, signal)
    } catch (error) {
      const kind = (error as { kind?: string })?.kind
      if (kind === 'timeout') fail('timeout')
      if (kind === 'cancelled') fail('cancelled')
      if (kind === 'stale-document' || kind === 'stale-selection' || kind === 'read-only')
        fail('document-changed')
      if (kind === 'closed' || kind === 'unavailable' || kind === 'disposed')
        fail('editor-replaced')
      if (kind === 'unsupported-read' || kind === 'unsafe-content') fail('unsupported-read')
      fail('editor-error')
    }
  }

  private stopForError(run: QueuedRun, error: unknown): void {
    const reason = error instanceof WorkspaceExecutionError ? error.kind : 'editor-error'
    const stopped = this.registry.stopRun(
      this.owner,
      run.binding.conversationId,
      run.binding.runId,
      run.binding.bindingId,
      reason,
    )
    if (!stopped.ok) this.finish(run)
  }
  private stopped(stopped: WorkspaceAiStoppedRun): void {
    const run = this.runs.get(stopped.bindingId)
    if (run) {
      run.abort.abort()
      this.finish(run)
    }
    this.deps.onStopped?.({ ...stopped })
  }
  private finish(run: QueuedRun): void {
    if (run.timer) clearTimeout(run.timer)
    if (this.active === run) this.active = null
    this.runs.delete(run.binding.bindingId)
    run.session = null
    if (run.target.resource.client.available) {
      const releasing = run.target.resource.client
        .request({ kind: 'settle', runId: run.editorRunId })
        .then(
          () => {},
          () => {},
        )
      this.draining = Promise.all([this.draining, releasing]).then(() => {})
    }
  }
  private remember<T>(map: Map<string, T>, id: string, entry: T, limit: number): void {
    map.set(id, entry)
    if (map.size > limit) map.delete(map.keys().next().value!)
  }
}
