import type { WorkspaceReadRequest } from '@genoffice/agent-core'
import type { WorkspaceConversationOrigin } from '../../../shared/workspace-ai-api'
import type { DocumentTabKind } from '../../../shared/tabs-api'
import type {
  WorkspaceCapture,
  WorkspaceReferenceCapture,
} from '../../../main/workspace-ai/execution-controller'
import { workspaceConversationKey } from './conversations'
import type { WorkspaceRunHost } from './run-host'

export interface WorkspacePanelDocument {
  id: string
  title: string
  kind: DocumentTabKind
  /** Display metadata only. Availability must come from the owning main host. */
  available: boolean
  active?: boolean
  pane?: 'left' | 'right' | 'top' | 'bottom'
}

export interface WorkspacePanelSource {
  origin: WorkspaceConversationOrigin
  title: string
  projectTitle?: string
}

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
}

export interface WorkspacePanelDeps {
  host: WorkspaceRunHost
  /** Calls describe only; selecting a target must not read its body or start a model. */
  capture(tabId: string): Promise<WorkspaceCapture>
  /** Only the opaque handle/binding comes back to the panel, never the reference body. */
  reference(
    captureId: string,
    scope: WorkspaceReadRequest['scope'],
  ): Promise<Pick<WorkspaceReferenceCapture, 'id' | 'binding'>>
  create(): Promise<WorkspacePanelSource>
}

/** A view owner around the existing run host. Hiding/unmounting its UI never disposes a run. */
export class WorkspacePanelController {
  readonly host: WorkspaceRunHost
  private documents: WorkspacePanelDocument[] = []
  private sources: WorkspacePanelSource[] = []
  private readonly selections = new Map<string, Selection>()
  private readonly targetEpochs = new Map<string, number>()
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
      value = { target: null, references: [], pending: 0, error: false, submitting: false }
      this.selections.set(id, value)
    }
    return structuredClone(value)
  }
  private change(id: string, patch: Partial<Selection>): void {
    this.selections.set(id, { ...this.selection(id), ...patch })
    this.notify()
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

  async create(): Promise<void> {
    if (this.closed) return
    const epoch = ++this.opening
    // Only a new conversation proposes the focused document, captured at this click.
    const initial = this.documents.find((item) => item.active && item.available)?.id
    this.openingFailed = false
    try {
      const source = await this.deps.create()
      if (this.closed) return
      this.sources.push(structuredClone(source))
      const id = await this.host.conversations.open(source.origin)
      if (this.closed) return
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

  async send(id: string): Promise<void> {
    if (!this.canSend(id)) return
    const selection = this.selection(id)
    const composer = this.host.conversations.state(id).composer
    this.change(id, { submitting: true, error: false })
    try {
      const run = await this.host.submit(
        id,
        selection.target!.id,
        selection.references.map((r) => r.id),
      )
      if (
        !this.closed &&
        run &&
        JSON.stringify(this.host.conversations.state(id).composer) === JSON.stringify(composer)
      ) {
        this.host.conversations.edit(id, {
          composer: { ...composer, text: '', attachments: [], skills: [] },
        })
      }
    } catch {
      if (!this.closed) this.change(id, { error: true })
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
