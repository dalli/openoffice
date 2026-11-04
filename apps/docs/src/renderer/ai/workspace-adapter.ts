import type { Editor } from '@tiptap/core'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import type { Selection, Transaction } from '@tiptap/pm/state'
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

/** Body-only capabilities; media, network, files and app-owned side stores stay outside this adapter. */
export const DOCS_WORKSPACE_TOOL_ACCESS = {
  get_document_context: 'read',
  read_blocks: 'read',
  read_revisions: 'read',
  insert_content: 'write',
  replace_blocks: 'write',
  replace_selection: 'write',
  apply_ops: 'write',
  accept_changes: 'write',
  reject_changes: 'write',
  insert_chart: 'write',
  edit_chart: 'write',
} as const

export interface DocsWorkspaceAdapterDeps {
  getEditor(): Editor | null
  /** Opaque document-load token; changes on reload/new file, even when the same Editor is reused. */
  getDocumentId(): string
  getNumIds(): NumIds
  getTrack?(): AiTrack | undefined
  beforeFirstWrite?(editor: Editor, signal?: AbortSignal): void | Promise<void>
}

interface SelectionSnapshot {
  readonly id: string
  readonly instanceId: string
  readonly revision: string
  readonly editor: Editor
  readonly doc: ProseMirrorNode
  readonly scope: SelectionScope
  readonly selection: Selection
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
  let attached: Editor | null = null
  let lastDoc: ProseMirrorNode | null = null
  let documentId: string | null = null
  let editorSequence = 0
  let revision = 0
  let selectionSequence = 0
  let prepareFirstWrite: ((signal?: AbortSignal) => void | Promise<void>) | null = null
  const selections = new Map<string, SelectionSnapshot>()
  const changes = new Set<() => void>()

  const onTransaction = ({ transaction }: { transaction: Transaction }): void => {
    if (!transaction.docChanged) return
    revision++
    lastDoc = attached?.state.doc ?? null
    selections.clear()
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
    const selection = state.selectionId ? selections.get(state.selectionId) : undefined
    if (
      !selection ||
      selection.editor !== live.editor ||
      selection.doc !== live.editor.state.doc ||
      selection.instanceId !== state.instanceId ||
      selection.revision !== state.revision
    ) {
      throw new WorkspaceAdapterError('stale-selection')
    }
    return selection
  }

  return createWorkspaceDocumentAdapter({
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
    createSkill(state) {
      const selection = selectionFor(state)
      let expectedRevision = selection.revision
      let expectedDoc = selection.doc
      let mappedSelection = selection.selection
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
      const original = createDocsSkill(
        () => selection.editor,
        () => deps.getNumIds(),
        () => deps.getTrack?.(),
        undefined,
        undefined,
        () => false,
        undefined,
        undefined,
        undefined,
        undefined,
        () => false,
        undefined,
        () => {
          assertCurrent()
          return getSelectionScope(selection.editor, mappedSelection)
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
        return result
      }
      return {
        id: original.id,
        get systemPrompt() {
          assertCurrent()
          return (
            original.systemPrompt +
            '\n\nThis workspace adapter exposes only local document-body tools. Use only the tools in the current list. Network, media, new-file creation, saving, comments, notes and header/footer stores are unavailable here.'
          )
        },
        get tools() {
          return original.tools.filter((tool) =>
            Object.hasOwn(DOCS_WORKSPACE_TOOL_ACCESS, tool.name),
          )
        },
        buildContext: () => {
          assertCurrent()
          return original.buildContext?.() ?? ''
        },
        executeTool(call, signal) {
          const access =
            DOCS_WORKSPACE_TOOL_ACCESS[call.name as keyof typeof DOCS_WORKSPACE_TOOL_ACCESS]
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
            // replace_selection deliberately selects the new text; formatting
            // and structural tools only map the original captured selection.
            mappedSelection =
              call.name === 'replace_selection' && transaction.selectionSet
                ? transaction.selection
                : mappedSelection.map(transaction.doc, transaction.mapping)
          }
          selection.editor.on('transaction', mapSelection)
          let result: ToolExecution | Promise<ToolExecution>
          try {
            result = original.executeTool(call, signal)
          } finally {
            selection.editor.off('transaction', mapSelection)
          }
          return result instanceof Promise ? result.then(settle) : settle(result)
        },
        verifyResponse: original.verifyResponse,
      } satisfies AgentSkill
    },
    toolAccess: DOCS_WORKSPACE_TOOL_ACCESS,
    rollback: {
      capture: () => captureRollbackSnapshot(current().editor),
      restore: (snapshot) => restoreRollbackSnapshot(current().editor, snapshot as ProseMirrorNode),
    },
    beforeFirstWrite: (signal) => {
      if (!prepareFirstWrite) throw new WorkspaceAdapterError('not-ready')
      return prepareFirstWrite(signal)
    },
    dispose() {
      attached?.off('transaction', onTransaction)
      attached = null
      lastDoc = null
      prepareFirstWrite = null
      selections.clear()
    },
  })
}
