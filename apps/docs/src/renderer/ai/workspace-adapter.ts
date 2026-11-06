import type { Editor } from '@tiptap/core'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import { Selection, type Transaction } from '@tiptap/pm/state'
import {
  createWorkspaceDocumentAdapter,
  WorkspaceAdapterError,
  type AgentSkill,
  type ToolExecution,
  type WorkspaceAdapterState,
  type WorkspaceDocumentAdapter,
} from '@genoffice/agent-core'
import { createDocsSkill } from './docs-skill'
import { getSelectionScope, type AiTrack, type NumIds, type SelectionScope } from './protocol'
import { isTrackedDeleted, liveText } from './doc-utils'
import { redactTextBetween } from './redact-view'
import { REDACT_MARK, sanitizeLabel } from './redact'
import { captureRollbackSnapshot, restoreRollbackSnapshot } from './rollback-snapshot'
import type { DocsReviewState, DocsReviewStore } from '../review-store'
import type { AiCommentsAccess, AiHeaderFooterAccess, AiDocExtras } from './tools'
import type { AiPageSetupAccess } from './page-setup'
import type { DocsLayoutState, DocsLayoutStore } from '../layout-store'
import type { AiNotesAccess } from './note-ops'
import { DRAFT_META } from './doc-writer'
import { WorkspaceDocWriter } from './workspace-writer'
import type { TemplateDocument } from '@genoffice/ui'

/** Body capabilities; additional tools require the corresponding observed App stores. */
export const DOCS_WORKSPACE_TOOL_ACCESS = {
  get_document_context: 'read',
  read_blocks: 'read',
  read_revisions: 'read',
  generate_mindmap: 'read',
  write_document: 'write',
  insert_content: 'write',
  replace_blocks: 'write',
  replace_selection: 'write',
  apply_ops: 'write',
  accept_changes: 'write',
  reject_changes: 'write',
  insert_chart: 'write',
  edit_chart: 'write',
} as const

const COMMENT_TOOL_ACCESS = {
  read_comments: 'read',
  add_comment: 'write',
  reply_comment: 'write',
  resolve_comment: 'write',
  delete_comment: 'write',
} as const

const NOTE_TOOL_ACCESS = {
  read_notes: 'read',
  insert_footnote: 'write',
  insert_endnote: 'write',
  edit_note: 'write',
  delete_note: 'write',
} as const

export interface DocsWorkspaceAdapterDeps {
  getEditor(): Editor | null
  /** Opaque document-load token; changes on reload/new file, even when the same Editor is reused. */
  getDocumentId(): string
  getNumIds(): NumIds
  getTrack?(): AiTrack | undefined
  layoutStore?: DocsLayoutStore
  getLayoutBaseline?(): unknown
  getHf?(): AiHeaderFooterAccess | undefined
  getPageSetup?(): AiPageSetupAccess | undefined
  getExtras?(): AiDocExtras | undefined
  reviewStore?: DocsReviewStore
  getComments?(): AiCommentsAccess | undefined
  getNotes?(): AiNotesAccess | undefined
  getTemplateDocument?(): TemplateDocument | null
  beforeFirstWrite?(editor: Editor, signal?: AbortSignal): void | Promise<void>
}

interface SelectionSnapshot {
  readonly id: string
  readonly instanceId: string
  revision: string
  readonly editor: Editor
  doc: ProseMirrorNode
  scope: SelectionScope
  selection: Selection
  valid: boolean
}

/** Bounded, in-memory selection handles. Eviction rejects an old request; it never substitutes a live range. */
const MAX_SELECTION_SNAPSHOTS = 128

function safeRedactions(doc: ProseMirrorNode): void {
  let redacted = false
  doc.descendants((node) => {
    for (const mark of node.marks) {
      if (mark.type.name !== REDACT_MARK) continue
      redacted = true
      const label: unknown = mark.attrs.label
      if (
        !node.isInline ||
        typeof label !== 'string' ||
        !label.trim() ||
        sanitizeLabel(label) !== label.trim()
      ) {
        throw new WorkspaceAdapterError('unsafe-content')
      }
    }
  })
  if (!redacted) return
  // The legacy context reader uses raw text for whole-block deletion previews.
  // Until that reader supports projection, this adapter refuses those reads.
  doc.forEach((block) => {
    if (!isTrackedDeleted(block)) return
    block.descendants((node) => {
      if (node.marks.some((mark) => mark.type.name === REDACT_MARK)) {
        throw new WorkspaceAdapterError('unsafe-content')
      }
    })
  })
}

/**
 * Keeps the existing DOCX executor in its renderer. There is no model call,
 * disk read, save, MCP read or editor focus change on this path.
 */
export function createDocsWorkspaceAdapter(
  deps: DocsWorkspaceAdapterDeps,
): WorkspaceDocumentAdapter {
  const writer = new WorkspaceDocWriter()
  let attached: Editor | null = null
  let lastDoc: ProseMirrorNode | null = null
  let documentId: string | null = null
  let editorSequence = 0
  let revision = 0
  let selectionSequence = 0
  let prepareFirstWrite: ((signal?: AbortSignal) => void | Promise<void>) | null = null
  const selections = new Map<string, SelectionSnapshot>()
  const changes = new Set<() => void>()
  let layoutBaseline = deps.getLayoutBaseline?.()
  let layoutRevision = deps.layoutStore?.revision ?? 0
  const offLayout = deps.layoutStore?.subscribe(() => {
    const next = deps.layoutStore!.revision
    if (next === layoutRevision) return
    layoutRevision = next
    revision++
    for (const notify of changes) notify()
  })
  let reviewRevision = deps.reviewStore?.revision ?? 0
  const offReview = deps.reviewStore?.subscribe(() => {
    const next = deps.reviewStore!.revision
    if (next === reviewRevision) return
    reviewRevision = next
    revision++
    for (const notify of changes) notify()
  })
  const toolAccess: Readonly<Record<string, 'read' | 'write'>> = {
    ...DOCS_WORKSPACE_TOOL_ACCESS,
    ...(deps.layoutStore && deps.getHf ? { set_header_footer: 'write' as const } : {}),
    ...(deps.layoutStore && deps.getPageSetup
      ? { set_page_setup: 'write' as const, insert_section_break: 'write' as const }
      : {}),
    ...(deps.layoutStore && deps.getExtras
      ? {
          list_styles: 'read' as const,
          define_style: 'write' as const,
          set_watermark: 'write' as const,
        }
      : {}),
    ...(deps.reviewStore && deps.getComments ? COMMENT_TOOL_ACCESS : {}),
    ...(deps.reviewStore && deps.getNotes ? NOTE_TOOL_ACCESS : {}),
  }

  const onTransaction = ({ transaction }: { transaction: Transaction }): void => {
    if (!transaction.docChanged) return
    lastDoc = attached?.state.doc ?? null
    if (transaction.getMeta(DRAFT_META)) return
    revision++
    for (const notify of changes) notify()
  }
  const current = (): { editor: Editor; instanceId: string; revision: string } => {
    const editor = deps.getEditor()
    const nextDocumentId = deps.getDocumentId()
    if (!editor || editor.isDestroyed || !nextDocumentId)
      throw new WorkspaceAdapterError('not-ready')
    if (attached !== editor) {
      attached?.off('transaction', onTransaction)
      attached = editor
      attached.on('transaction', onTransaction)
      editorSequence++
      lastDoc = editor.state.doc
      selections.clear()
    }
    if (documentId !== nextDocumentId || lastDoc !== editor.state.doc) {
      revision++
      documentId = nextDocumentId
      lastDoc = editor.state.doc
      selections.clear()
    }
    const nextBaseline = deps.getLayoutBaseline?.()
    if (deps.layoutStore && nextBaseline == null) throw new WorkspaceAdapterError('not-ready')
    if (layoutBaseline !== nextBaseline) {
      layoutBaseline = nextBaseline
      revision++
      selections.clear()
    }
    safeRedactions(editor.state.doc)
    return {
      editor,
      instanceId: `docs-${editorSequence}:${nextDocumentId}`,
      revision: String(revision),
    }
  }
  const selectionFor = (state: WorkspaceAdapterState): SelectionSnapshot => {
    const live = current()
    if (live.instanceId !== state.instanceId || live.revision !== state.revision) {
      throw new WorkspaceAdapterError('stale-document')
    }
    if (state.selectionId === null) {
      // A document-scoped request never borrows the user's current selection.
      const selection = Selection.atEnd(live.editor.state.doc)
      return {
        id: '',
        instanceId: live.instanceId,
        revision: live.revision,
        editor: live.editor,
        doc: live.editor.state.doc,
        scope: getSelectionScope(live.editor, selection),
        selection,
        valid: true,
      }
    }
    const selection = state.selectionId ? selections.get(state.selectionId) : undefined
    if (
      !selection ||
      !selection.valid ||
      selection.editor !== live.editor ||
      selection.doc !== live.editor.state.doc ||
      selection.instanceId !== state.instanceId ||
      selection.revision !== state.revision
    ) {
      throw new WorkspaceAdapterError('stale-selection')
    }
    return selection
  }

  const adapter = createWorkspaceDocumentAdapter({
    masking: 'applied',
    subscribeChanges(handler) {
      changes.add(handler)
      return () => {
        changes.delete(handler)
      }
    },
    getState() {
      const live = current()
      const scope = getSelectionScope(live.editor)
      let selection = [...selections.values()].find(
        (snapshot) =>
          snapshot.editor === live.editor &&
          snapshot.valid &&
          snapshot.doc === live.editor.state.doc &&
          snapshot.selection.eq(live.editor.state.selection),
      )
      if (!selection) {
        selection = {
          id: `docs-selection-${++selectionSequence}`,
          instanceId: live.instanceId,
          revision: live.revision,
          editor: live.editor,
          doc: live.editor.state.doc,
          scope: { ...scope },
          selection: live.editor.state.selection,
          valid: true,
        }
        selections.set(selection.id, selection)
        if (selections.size > MAX_SELECTION_SNAPSHOTS)
          selections.delete(selections.keys().next().value!)
      }
      return {
        instanceId: live.instanceId,
        revision: live.revision,
        selectionId: selection.id,
        readOnly: !live.editor.isEditable,
      }
    },
    refreshSelection(state, now) {
      const selection = state.selectionId ? selections.get(state.selectionId) : undefined
      if (
        !selection ||
        !selection.valid ||
        selection.instanceId !== now.instanceId ||
        selection.revision !== now.revision ||
        selection.doc !== current().editor.state.doc
      )
        throw new WorkspaceAdapterError('stale-selection')
      return selection.id
    },
    read(state, request) {
      const selection = selectionFor(state)
      if (request.scope === 'selection' && !selection.scope.isRange) {
        throw new WorkspaceAdapterError('unsupported-read')
      }
      const text =
        request.scope === 'selection'
          ? redactTextBetween(
              selection.editor.getJSON(),
              selection.scope.from!,
              selection.scope.to!,
              '\n',
            )
          : liveText(selection.editor.state.doc, false, { blockSeparator: '\n' })
      return { text, truncated: false }
    },
    async template(_state, mode, signal) {
      const source = deps.getTemplateDocument?.()
      if (!source) return null
      const hidden = () => {
        let redacted = false
        current().editor.state.doc.descendants((node) => {
          if (node.marks.some((mark) => mark.type.name === REDACT_MARK)) redacted = true
        })
        return redacted
      }
      const redacted = hidden() || (await source.redacted()) !== false
      if (mode === 'as-is' && redacted) throw new WorkspaceAdapterError('unsafe-content')
      const bytes = mode === 'info' ? null : await source.bytes()
      if (signal?.aborted) throw new WorkspaceAdapterError('cancelled')
      return {
        fileName: source.fileName,
        title: source.title,
        redacted: redacted || hidden(),
        bytes,
      }
    },
    createSkill(state) {
      const selection = selectionFor(state)
      let expectedRevision = selection.revision
      let expectedDoc = selection.doc
      let mappedSelection = selection.selection
      const scope = (): SelectionScope =>
        state.selectionId === null
          ? {
              startIndex: expectedDoc.childCount - 1,
              endIndex: expectedDoc.childCount - 1,
              isRange: false,
            }
          : getSelectionScope(selection.editor, mappedSelection)
      const assertCurrent = (signal?: AbortSignal, writing = false): void => {
        if (signal?.aborted) throw new WorkspaceAdapterError('cancelled')
        const live = current()
        if (
          live.editor !== selection.editor ||
          live.instanceId !== selection.instanceId ||
          live.revision !== expectedRevision ||
          live.editor.state.doc !== expectedDoc
        ) {
          throw new WorkspaceAdapterError('stale-document')
        }
        if (writing && !live.editor.isEditable) throw new WorkspaceAdapterError('read-only')
      }
      prepareFirstWrite = (signal) => {
        assertCurrent(signal, true)
        return deps.beforeFirstWrite?.(selection.editor, signal)
      }
      let activeWriter: import('./doc-writer').AiDocWriter | undefined
      const original = createDocsSkill(
        () => selection.editor,
        () => deps.getNumIds(),
        () => deps.getTrack?.(),
        deps.reviewStore ? deps.getComments : undefined,
        deps.layoutStore ? deps.getHf : undefined,
        () => false,
        () => activeWriter,
        deps.layoutStore ? deps.getPageSetup : undefined,
        deps.layoutStore ? deps.getExtras : undefined,
        deps.reviewStore ? deps.getNotes : undefined,
        () => false,
        undefined,
        () => {
          assertCurrent()
          return scope()
        },
      )
      const settle = (result: ToolExecution): ToolExecution => {
        const live = current()
        if (live.instanceId !== selection.instanceId || live.editor !== selection.editor) {
          throw new WorkspaceAdapterError('stale-document')
        }
        if (
          !result.mutated &&
          (live.revision !== expectedRevision || live.editor.state.doc !== expectedDoc)
        ) {
          throw new WorkspaceAdapterError('stale-document')
        }
        expectedRevision = live.revision
        expectedDoc = live.editor.state.doc
        selection.revision = live.revision
        selection.doc = expectedDoc
        selection.selection = mappedSelection
        selection.scope = scope()
        return result
      }
      return {
        id: original.id,
        get systemPrompt() {
          assertCurrent()
          return (
            original.systemPrompt +
            '\n\nThis workspace adapter exposes only the listed local document tools. Network, media, new-file creation and saving are unavailable here. Long-form writing uses write_document when listed. Header/footer, page setup, style, watermark, comment and note tools are available only when listed.'
          )
        },
        get tools() {
          return original.tools
            .filter((tool) => Object.hasOwn(toolAccess, tool.name))
            .map((tool) =>
              tool.name === 'generate_mindmap'
                ? {
                    ...tool,
                    description: tool.description.replace(
                      ' so the user can jump to the passage',
                      '',
                    ),
                  }
                : tool,
            )
        },
        buildContext: () => {
          assertCurrent()
          return original.buildContext?.() ?? ''
        },
        executeTool(call, signal) {
          const access = toolAccess[call.name]
          if (!access) throw new WorkspaceAdapterError('unknown-tool')
          assertCurrent(signal, access === 'write')
          if (access === 'write' && deps.getTrack?.()) {
            // A tracked rewrite could create an unsafe legacy deletion preview
            // after applying. Refuse before mutation rather than fail afterwards.
            selection.editor.state.doc.descendants((node) => {
              if (node.marks.some((mark) => mark.type.name === REDACT_MARK)) {
                throw new WorkspaceAdapterError('unsafe-content')
              }
            })
          }
          const mapSelection = ({ transaction }: { transaction: Transaction }): void => {
            if (!transaction.docChanged) return
            if (transaction.getMeta(DRAFT_META)) {
              expectedDoc = selection.editor.state.doc
              return
            }
            const previous = mappedSelection
            // replace_selection deliberately selects the new text; formatting
            // and structural tools only map the original captured selection.
            mappedSelection =
              call.name === 'replace_selection' && transaction.selectionSet
                ? transaction.selection
                : mappedSelection.map(transaction.doc, transaction.mapping)
            if (
              (!previous.empty && mappedSelection.empty) ||
              (previous.empty && transaction.mapping.mapResult(previous.head, 1).deletedAcross)
            )
              selection.valid = false
          }
          selection.editor.on('transaction', mapSelection)
          const beforeReview = reviewRevision
          const beforeLayout = layoutRevision
          let result: ToolExecution | Promise<ToolExecution>
          const cleanup = () => {
            selection.editor.off('transaction', mapSelection)
            if (call.name === 'write_document') {
              writer.finish()
              activeWriter = undefined
            }
          }
          try {
            if (call.name === 'write_document') {
              const remote = writer.begin(call.id)
              activeWriter = {
                async write(spec, progress, writeSignal) {
                  const result = await remote.write(
                    spec,
                    (html) => {
                      assertCurrent(writeSignal, true)
                      progress(html)
                    },
                    writeSignal,
                  )
                  assertCurrent(writeSignal, true)
                  return result
                },
              }
            }
            result = original.executeTool(call, signal)
          } catch (error) {
            cleanup()
            throw error
          }

          const settleStores = (execution: ToolExecution): ToolExecution =>
            settle({
              ...execution,
              ...(execution.display?.kind === 'mindmap'
                ? {
                    output:
                      'Mind map rendered as an interactive workspace chat card. The user can expand branches, export PNG/SVG, or fill a follow-up question for the same target document. This does not modify the document. Reply with one short sentence; do not repeat the map content.',
                  }
                : {}),
              mutated:
                execution.mutated ||
                beforeReview !== reviewRevision ||
                beforeLayout !== layoutRevision,
            })
          if (result instanceof Promise) return result.then(settleStores).finally(cleanup)
          try {
            return settleStores(result)
          } finally {
            cleanup()
          }
        },
        verifyResponse: original.verifyResponse,
      } satisfies AgentSkill
    },
    toolAccess,
    rollback: {
      capture: (state) => ({
        doc: captureRollbackSnapshot(current().editor),
        review: deps.reviewStore?.capture(),
        layout: deps.layoutStore?.capture(),
        selections: [...selections.values()]
          .filter((selection) => selection.valid && selection.revision === state.revision)
          .map((selection) => ({ ...selection })),
      }),
      restore: (snapshot) => {
        const saved = snapshot as {
          doc: ProseMirrorNode
          review?: DocsReviewState
          layout?: DocsLayoutState
          selections: SelectionSnapshot[]
        }
        const editor = current().editor
        if (!editor.state.doc.eq(saved.doc)) restoreRollbackSnapshot(editor, saved.doc)
        if (saved.review) deps.reviewStore!.restore(saved.review)
        if (saved.layout) deps.layoutStore!.restore(saved.layout)
        const live = current()
        for (const selection of saved.selections) {
          if (selection.instanceId === live.instanceId && selection.doc.eq(live.editor.state.doc))
            selections.set(selection.id, {
              ...selection,
              revision: live.revision,
              doc: live.editor.state.doc,
              selection: selection.selection.getBookmark().resolve(live.editor.state.doc),
            })
        }
        while (selections.size > MAX_SELECTION_SNAPSHOTS)
          selections.delete(selections.keys().next().value!)
      },
    },
    beforeFirstWrite: (signal) => {
      if (!prepareFirstWrite) throw new WorkspaceAdapterError('not-ready')
      return prepareFirstWrite(signal)
    },
    dispose() {
      offReview?.()
      offLayout?.()
      attached?.off('transaction', onTransaction)
      attached = null
      lastDoc = null
      prepareFirstWrite = null
      selections.clear()
    },
  })
  return { ...adapter, updateWriter: (id, update, signal) => writer.update(id, update, signal) }
}
