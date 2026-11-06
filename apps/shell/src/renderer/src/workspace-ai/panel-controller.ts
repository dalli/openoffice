import type { WorkspaceReadRequest } from '@genoffice/agent-core'
import type {
  WorkspacePanelDocument,
  WorkspacePanelSource,
  WorkspacePanelEntry,
} from '../../../shared/workspace-panel-api'
export type {
  WorkspacePanelDocument,
  WorkspacePanelSource,
} from '../../../shared/workspace-panel-api'
import type {
  WorkspaceCapture,
  WorkspaceReferenceCapture,
} from '../../../main/workspace-ai/execution-controller'
import { workspaceConversationKey } from './conversations'
import type { WorkspaceRunHost } from './run-host'
import type { ProjectNotice } from '@genoffice/project-store'
import type { WorkspaceDocumentIdentity } from '../../../shared/workspace-ai-api'

interface Selection {
  target: WorkspaceCapture | null
  references: Array<{
    id: string
    binding: WorkspaceReferenceCapture['binding']
    scope: WorkspaceReadRequest['scope']
    title: string
  }>
  pending: number
  error: boolean
  submitting: boolean
  projectNotices: ProjectNotice[]
}

export interface WorkspacePanelDeps {
  host: WorkspaceRunHost
  /** Calls describe only; selecting a target must not read its body or start a model. */
  capture(tabId: string): Promise<WorkspaceCapture>
  /** Metadata-only renewal of the pinned target; never captures a new visible selection. */
  refresh(captureId: string): Promise<WorkspaceCapture>
  /** Only the opaque handle/binding comes back to the panel, never the reference body. */
  reference(
    captureId: string,
    scope: WorkspaceReadRequest['scope'],
  ): Promise<Pick<WorkspaceReferenceCapture, 'id' | 'binding'>>
  create(projectId?: string): Promise<WorkspacePanelSource>
  handoff?(
    id: string,
    capture: WorkspaceCapture,
    entry: Extract<WorkspacePanelEntry, { kind: 'template-handoff' }>,
  ): boolean
}

/** A view owner around the existing run host. Hiding/unmounting its UI never disposes a run. */
export class WorkspacePanelController {
  readonly host: WorkspaceRunHost
  private documents: WorkspacePanelDocument[] = []
  private sources: WorkspacePanelSource[] = []
  private readonly selections = new Map<string, Selection>()
  private readonly targetEpochs = new Map<string, number>()
  private readonly batches = new Map<
    string,
    { captureId: string; text: string; displayText: string }
  >()
  private readonly handoffs = new Map<string, { captureId: string; text: string; chips: string }>()
  private readonly referenceEpochs = new Map<string, number>()
  private readonly listeners = new Set<() => void>()
  private readonly off: () => void
  private revision = 0
  private opening = 0
  private closed = false
  private openingFailed = false

  constructor(private readonly deps: WorkspacePanelDeps) {
    this.host = deps.host
    this.off = this.host.subscribe(() => this.notify())
  }

  subscribe = (listener: () => void): (() => void) => {
    this.listeners.add(listener)
    return () => this.listeners.delete(listener)
  }
  version = (): number => this.revision
  private notify(): void {
    if (this.closed) return
    this.revision++
    for (const listener of [...this.listeners]) listener()
  }

  /** Catalog updates do not capture, retarget, load recovery tabs or read any history. */
  catalog(
    documents: readonly WorkspacePanelDocument[],
    sources: readonly WorkspacePanelSource[],
  ): void {
    this.documents = structuredClone([...documents])
    this.sources = structuredClone([...sources])
    this.notify()
  }
  documentList(): WorkspacePanelDocument[] {
    return structuredClone(this.documents)
  }
  sourceList(): WorkspacePanelSource[] {
    return structuredClone(this.sources)
  }
  hasOpenError(): boolean {
    return this.openingFailed
  }
  selection(id: string): Selection {
    let value = this.selections.get(id)
    if (!value) {
      value = {
        target: null,
        references: [],
        pending: 0,
        error: false,
        submitting: false,
        projectNotices: [],
      }
      this.selections.set(id, value)
    }
    return structuredClone(value)
  }
  private change(id: string, patch: Partial<Selection>): void {
    this.selections.set(id, { ...this.selection(id), ...patch })
    this.notify()
  }
  projectNotices(id: string, projectNotices: ProjectNotice[]): void {
    if (!this.closed)
      this.change(id, {
        projectNotices: [
          ...this.selection(id).projectNotices.filter(
            (notice) =>
              notice.kind === 'sync-folder' &&
              !projectNotices.some((next) => next.kind === 'sync-folder'),
          ),
          ...projectNotices,
        ],
      })
  }

  async open(source: WorkspacePanelSource): Promise<void> {
    if (this.closed) return
    const epoch = ++this.opening
    this.openingFailed = false
    try {
      const id = await this.host.conversations.open(source.origin)
      if (this.closed || epoch !== this.opening) return
      this.host.conversations.select(id)
    } catch {
      if (!this.closed && epoch === this.opening) this.openingFailed = true
    }
    this.notify()
  }

  select(id: string): void {
    ++this.opening
    this.openingFailed = false
    this.host.conversations.select(id)
  }

  async create(projectId?: string, notices: ProjectNotice[] = []): Promise<void> {
    if (this.closed) return
    const epoch = ++this.opening
    // Only a new conversation proposes the focused document, captured at this click.
    const initial = this.documents.find((item) => item.active && item.available)?.id
    this.openingFailed = false
    try {
      const source = await this.deps.create(projectId)
      if (this.closed) return
      if (
        !this.sources.some(
          (item) =>
            workspaceConversationKey(item.origin) === workspaceConversationKey(source.origin),
        )
      )
        this.sources.push(structuredClone(source))
      const id = await this.host.conversations.open(source.origin)
      if (this.closed) return
      this.projectNotices(id, notices)
      if (epoch === this.opening) this.host.conversations.select(id)
      if (initial) await this.target(id, initial)
    } catch {
      if (!this.closed && epoch === this.opening) this.openingFailed = true
    }
    this.notify()
  }

  private available(tabId: string): boolean {
    return this.documents.some((item) => item.id === tabId && item.available)
  }

  /** Explicit editor actions own a new conversation; existing drafts and targets stay intact. */
  async enter(
    capture: WorkspaceCapture,
    entry: WorkspacePanelEntry,
    acceptBatch?: (accepted?: boolean, conversationId?: string) => Promise<boolean>,
  ): Promise<void> {
    if (this.closed) return
    if (entry.kind === 'queue') {
      const selected = this.host.conversations.selectedId
      const target = selected ? this.selection(selected).target?.binding.document : null
      if (
        target &&
        target.tabId === capture.binding.document.tabId &&
        target.documentId === capture.binding.document.documentId &&
        target.generation === capture.binding.document.generation
      )
        return
    }
    if (
      entry.kind !== 'preset' &&
      entry.kind !== 'edit-batch' &&
      entry.kind !== 'template-handoff' &&
      entry.kind !== 'queue' &&
      this.host.conversations.selectedId
    )
      return
    const epoch = ++this.opening
    this.openingFailed = false
    let enteredId: string | null = null
    try {
      const source = await this.deps.create(
        entry.kind === 'template-handoff' ? entry.projectId : undefined,
      )
      if (this.closed) return
      this.sources.push(structuredClone(source))
      const id = await this.host.conversations.open(source.origin)
      if (this.closed) return
      enteredId = id
      const composer = this.host.conversations.state(id).composer
      this.host.conversations.edit(id, {
        composer: {
          ...composer,
          target: capture.binding,
          text:
            entry.kind === 'preset' || entry.kind === 'template-handoff'
              ? entry.text
              : entry.kind === 'edit-batch'
                ? entry.displayText
                : '',
          ...(entry.kind === 'template-handoff' ? { skills: structuredClone(entry.chips) } : {}),
        },
      })
      this.change(id, { target: capture, error: !this.available(capture.binding.document.tabId) })
      if (epoch === this.opening) this.host.conversations.select(id)
      if (entry.kind === 'template-handoff') {
        this.change(id, { submitting: true })
        const supported = this.deps.handoff?.(id, capture, entry) === true
        if (!acceptBatch || !(await acceptBatch(supported, id)) || !supported || this.closed) {
          if (!this.closed) {
            const now = this.host.conversations.state(id).composer
            if (
              now.text === entry.text &&
              JSON.stringify(now.skills) === JSON.stringify(entry.chips)
            )
              this.host.conversations.edit(id, { composer: { ...now, text: '', skills: [] } })
            this.change(id, { error: true, submitting: false })
          }
          return
        }
        this.handoffs.set(id, {
          captureId: capture.id,
          text: entry.text,
          chips: JSON.stringify(entry.chips),
        })
        this.change(id, { submitting: false })
        const now = this.host.conversations.state(id).composer
        if (
          now.text === entry.text &&
          JSON.stringify(now.skills) === JSON.stringify(entry.chips) &&
          this.selection(id).target?.id === capture.id
        )
          await this.send(id)
      }
      if (entry.kind === 'edit-batch') {
        this.batches.set(id, {
          captureId: capture.id,
          text: entry.text,
          displayText: entry.displayText,
        })
        this.change(id, { submitting: true })
        if (!acceptBatch || !(await acceptBatch()) || this.closed) {
          this.batches.delete(id)
          if (!this.closed) {
            const now = this.host.conversations.state(id).composer
            if (now.text === entry.displayText)
              this.host.conversations.edit(id, { composer: { ...now, text: '' } })
            this.change(id, { error: true, submitting: false })
          }
          return
        }
        this.change(id, { submitting: false })
        if (
          this.host.conversations.state(id).composer.text === entry.displayText &&
          this.selection(id).target?.id === capture.id
        )
          await this.send(id)
      }
      if (entry.kind === 'preset' && entry.autoRun) await this.send(id)
    } catch {
      if (
        !this.closed &&
        enteredId &&
        (entry.kind === 'edit-batch' || entry.kind === 'template-handoff')
      ) {
        this.batches.delete(enteredId)
        this.handoffs.delete(enteredId)
        if (entry.kind === 'template-handoff') {
          // An interrupted acknowledgement may already have transferred ownership. Keep the draft,
          // block automatic sending and never replay the request in the editor.
        }
        this.change(enteredId, { error: true, submitting: false })
      }
      if (entry.kind === 'template-handoff') await acceptBatch?.(false).catch(() => {})
      if (!this.closed && epoch === this.opening) this.openingFailed = true
    }
    this.notify()
  }

  async target(id: string, tabId: string): Promise<void> {
    if (this.closed) return
    const epoch = (this.targetEpochs.get(id) ?? 0) + 1
    this.targetEpochs.set(id, epoch)
    // A failed retarget must never leave the previous document ready to receive a send.
    const previous = this.host.conversations.state(id).composer
    this.host.conversations.edit(id, { composer: { ...previous, target: null } })
    this.change(id, { target: null, pending: this.selection(id).pending + 1, error: false })
    try {
      if (!this.available(tabId)) throw new Error('Document unavailable')
      const target = await this.deps.capture(tabId)
      if (this.closed || this.targetEpochs.get(id) !== epoch) return
      if (!this.available(tabId) || target.binding.document.tabId !== tabId)
        throw new Error('Document unavailable')
      const composer = this.host.conversations.state(id).composer
      this.host.conversations.edit(id, { composer: { ...composer, target: target.binding } })
      this.change(id, { target })
    } catch {
      if (!this.closed && this.targetEpochs.get(id) === epoch) this.change(id, { error: true })
    } finally {
      if (!this.closed) this.change(id, { pending: this.selection(id).pending - 1 })
    }
  }

  async addReference(
    id: string,
    tabId: string,
    scope: WorkspaceReadRequest['scope'],
  ): Promise<void> {
    if (this.closed) return
    const key = JSON.stringify([id, tabId, scope])
    const epoch = (this.referenceEpochs.get(key) ?? 0) + 1
    this.referenceEpochs.set(key, epoch)
    this.change(id, { pending: this.selection(id).pending + 1, error: false })
    try {
      if (!this.available(tabId)) throw new Error('Document unavailable')
      const capture = await this.deps.capture(tabId)
      if (this.closed || this.referenceEpochs.get(key) !== epoch) return
      if (!this.available(tabId) || capture.binding.document.tabId !== tabId)
        throw new Error('Document unavailable')
      const reference = await this.deps.reference(capture.id, scope)
      if (this.closed || this.referenceEpochs.get(key) !== epoch) return
      if (!this.available(tabId)) throw new Error('Document unavailable')
      const references = this.selection(id).references.filter(
        (ref) => ref.binding.document.tabId !== tabId || ref.scope !== scope,
      )
      references.push({
        ...reference,
        scope,
        title: this.documents.find((d) => d.id === tabId)!.title,
      })
      const composer = this.host.conversations.state(id).composer
      this.host.conversations.edit(id, {
        composer: { ...composer, references: references.map((ref) => ref.binding) },
      })
      this.change(id, { references })
    } catch {
      if (!this.closed && this.referenceEpochs.get(key) === epoch) this.change(id, { error: true })
    } finally {
      if (!this.closed) this.change(id, { pending: this.selection(id).pending - 1 })
    }
  }

  removeReference(id: string, referenceId: string): void {
    const removed = this.selection(id).references.find((ref) => ref.id === referenceId)
    if (removed) {
      const key = JSON.stringify([id, removed.binding.document.tabId, removed.scope])
      this.referenceEpochs.set(key, (this.referenceEpochs.get(key) ?? 0) + 1)
    }
    const references = this.selection(id).references.filter((ref) => ref.id !== referenceId)
    const composer = this.host.conversations.state(id).composer
    this.host.conversations.edit(id, {
      composer: { ...composer, references: references.map((ref) => ref.binding) },
    })
    this.change(id, { references })
  }

  canSend(id: string): boolean {
    const value = this.selection(id)
    return (
      !!value.target &&
      !this.closed &&
      !value.error &&
      !value.pending &&
      !value.submitting &&
      this.available(value.target.binding.document.tabId) &&
      value.references.every((ref) => this.available(ref.binding.document.tabId))
    )
  }

  canPreset(id: string, document?: WorkspaceDocumentIdentity): boolean {
    if (!this.canSend(id)) return false
    const target = this.selection(id).target!.binding.document
    return (
      target.kind === 'docs' &&
      (!document ||
        (target.windowId === document.windowId &&
          target.tabId === document.tabId &&
          target.documentId === document.documentId &&
          target.generation === document.generation &&
          target.kind === document.kind))
    )
  }

  /** Keep the pinned target/selection and all other composer state. No capture or model call. */
  preset(id: string, text: string, document?: WorkspaceDocumentIdentity): boolean {
    if (!this.canPreset(id, document)) return false
    const composer = this.host.conversations.state(id).composer
    this.host.conversations.edit(id, { composer: { ...composer, text } })
    return true
  }

  async send(id: string): Promise<void> {
    if (!this.canSend(id)) return
    const selection = this.selection(id)
    const composer = this.host.conversations.state(id).composer
    const batch = this.batches.get(id)
    const sendingBatch = batch && composer.text === batch.displayText ? batch : null
    const handoff = this.handoffs.get(id)
    const sendingHandoff =
      handoff && composer.text === handoff.text && JSON.stringify(composer.skills) === handoff.chips
        ? handoff
        : null
    let targetId = selection.target!.id
    this.change(id, { submitting: true, error: false })
    try {
      if (sendingBatch && sendingBatch.captureId !== selection.target!.id)
        throw new Error('Edit batch target changed')
      if (sendingHandoff && sendingHandoff.captureId !== selection.target!.id)
        throw new Error('Template handoff target changed')
      // Batch block indexes are frozen at handoff. Never renew them onto a newer document.
      const target =
        sendingBatch || sendingHandoff
          ? selection.target!
          : await this.deps.refresh(selection.target!.id)
      if (this.closed) return
      const current = this.selection(id)
      if (
        current.target?.id !== selection.target!.id ||
        JSON.stringify(current.references) !== JSON.stringify(selection.references)
      )
        return
      const input = { ...composer, target: target.binding }
      targetId = target.id
      this.change(id, { target })
      this.host.conversations.edit(id, {
        composer: { ...this.host.conversations.state(id).composer, target: target.binding },
      })
      const run = await this.host.resume({
        conversationId: id,
        targetId: target.id,
        referenceIds: selection.references.map((r) => r.id),
        input: sendingBatch ? { ...input, text: sendingBatch.text } : input,
        ...(sendingBatch ? { displayText: sendingBatch.displayText } : {}),
      })
      if (
        !this.closed &&
        run &&
        !this.host.conversations.state(id).error &&
        JSON.stringify(this.host.conversations.state(id).composer) === JSON.stringify(input)
      ) {
        this.host.conversations.edit(id, {
          composer: { ...input, text: '', attachments: [], skills: [] },
        })
      }
    } catch {
      if (!this.closed && this.selection(id).target?.id === targetId)
        this.change(id, { error: true })
    } finally {
      if (!this.closed) this.change(id, { submitting: false })
    }
  }

  sourceFor(id: string): WorkspacePanelSource | undefined {
    const origin = this.host.conversations.state(id).origin
    return this.sourceList().find(
      (source) => workspaceConversationKey(source.origin) === workspaceConversationKey(origin),
    )
  }

  /** The window's owner disposes the run host separately, after stopping its transports/editors. */
  dispose(): void {
    this.closed = true
    this.off()
    this.listeners.clear()
  }
}
