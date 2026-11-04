import type { Editor, JSONContent } from '@tiptap/core'
import { TextSelection, type Selection, type Transaction } from '@tiptap/pm/state'
import type { Node as PmNode } from '@tiptap/pm/model'
import {
  createWorkspaceDocumentAdapter,
  WorkspaceAdapterError,
  type AgentSkill,
  type ToolExecution,
  type WorkspaceAdapterState,
  type WorkspaceDocumentAdapter,
  type WorkspaceDocumentAdapterDeps,
} from '@genoffice/agent-core'
import { createMarkdownSkill } from './markdown-skill'
import { frontmatterHasPrivateMarkup } from './frontmatter-model'
import { redactJson } from '../editor/redact'
import { validateOps, type FrontmatterAccess } from '../editor/ops'

export interface MarkdownWorkspaceAdapterDeps {
  rollback?: WorkspaceDocumentAdapterDeps['rollback']
  getEditor(): Editor | null
  /** Opaque document/load identity, changed when a file is replaced in the same editor. */
  getDocumentId(): string
  /** Also true for the Markdown source-view buffer while the block editor is still mounted. */
  isSourceMode(): boolean
  frontmatter?: FrontmatterAccess
  beforeFirstWrite?(signal?: AbortSignal): void | Promise<void>
  /** Frontmatter, source mode and document-load changes outside PM transactions. */
  subscribeChanges?(handler: () => void): () => void
}

const TOOL_ACCESS = {
  get_document_context: 'read',
  read_blocks: 'read',
  read_frontmatter: 'read',
  apply_ops: 'write',
} as const
const SNAPSHOT_LIMIT = 128

interface FrozenSelection {
  id: string
  instanceId: string
  revision: string
  selection: Selection
  valid: boolean
}

function reject(kind: ConstructorParameters<typeof WorkspaceAdapterError>[0]): never {
  throw new WorkspaceAdapterError(kind)
}

/** Preserve the selected range across our own edits instead of following a later click. */
function mapSelection(selection: Selection, transaction: Transaction): Selection {
  if (!(selection instanceof TextSelection) || selection.empty) {
    return selection.getBookmark().map(transaction.mapping).resolve(transaction.doc)
  }
  const from = transaction.mapping.map(selection.from, -1)
  const to = transaction.mapping.map(selection.to, 1)
  return TextSelection.between(
    transaction.doc.resolve(selection.anchor <= selection.head ? from : to),
    transaction.doc.resolve(selection.anchor <= selection.head ? to : from),
  )
}

function containsImage(node: JSONContent): boolean {
  return node.type === 'image' || node.content?.some(containsImage) === true
}

/**
 * The renderer keeps its real editor, redaction, operations and undo history.
 * This boundary has no disk, network, writer or model capability.
 */
export function createMarkdownWorkspaceAdapter(
  deps: MarkdownWorkspaceAdapterDeps,
): WorkspaceDocumentAdapter {
  let editor: Editor | null = null
  let documentId = ''
  let instance = 0
  let revision = 0
  let seenDoc: PmNode | null = null
  let seenFrontmatter = ''
  let selectionSequence = 0
  const selections = new Map<string, FrozenSelection>()
  const changes = new Set<() => void>()

  const changed = ({
    transaction,
    appendedTransactions = [],
  }: {
    transaction: Transaction
    appendedTransactions?: Transaction[]
  }) => {
    if ([transaction, ...appendedTransactions].some((item) => item.docChanged)) {
      revision++
      seenDoc = editor?.state.doc ?? null
      for (const notify of changes) notify()
    }
  }
  const syncRevision = (live: Editor) => {
    const frontmatter = deps.frontmatter?.read() ?? ''
    if (seenDoc !== live.state.doc || seenFrontmatter !== frontmatter) {
      revision++
      seenDoc = live.state.doc
      seenFrontmatter = frontmatter
    }
  }
  const liveEditor = (): Editor => {
    const live = deps.getEditor()
    const id = deps.getDocumentId()
    if (deps.isSourceMode() || !live || live.isDestroyed || !id) reject('not-ready')
    if (live !== editor || id !== documentId) {
      editor?.off('transaction', changed)
      editor = live
      documentId = id
      instance++
      revision = 1
      seenDoc = live.state.doc
      seenFrontmatter = deps.frontmatter?.read() ?? ''
      selections.clear()
      editor.on('transaction', changed)
    }
    syncRevision(live)
    return live
  }
  const instanceId = () => `markdown:${documentId}:${instance}`
  const state = (): WorkspaceAdapterState => {
    const live = liveEditor()
    const id = instanceId()
    const version = String(revision)
    let frozen = [...selections.values()].find(
      (candidate) =>
        candidate.valid &&
        candidate.instanceId === id &&
        candidate.revision === version &&
        candidate.selection.eq(live.state.selection),
    )
    if (!frozen) {
      frozen = {
        id: `markdown-selection-${++selectionSequence}`,
        instanceId: id,
        revision: version,
        selection: live.state.selection,
        valid: true,
      }
      selections.set(frozen.id, frozen)
      if (selections.size > SNAPSHOT_LIMIT) selections.delete(selections.keys().next().value!)
    }
    return { instanceId: id, revision: version, selectionId: frozen.id, readOnly: !live.isEditable }
  }
  const frozenFor = (captured: WorkspaceAdapterState): FrozenSelection => {
    const frozen = captured.selectionId ? selections.get(captured.selectionId) : undefined
    if (
      !frozen ||
      frozen.instanceId !== captured.instanceId ||
      frozen.revision !== captured.revision
    ) {
      reject('stale-selection')
    }
    return frozen
  }
  const safeFrontmatter = (): string => {
    const raw = deps.frontmatter?.read() ?? ''
    if (frontmatterHasPrivateMarkup(raw)) reject('unsafe-content')
    return raw
  }
  const assertReady = (
    expectedEditor: Editor,
    expectedRevision: string,
    expectedId: string,
    signal?: AbortSignal,
  ) => {
    if (signal?.aborted) reject('cancelled')
    const live = liveEditor()
    if (
      live !== expectedEditor ||
      instanceId() !== expectedId ||
      String(revision) !== expectedRevision
    )
      reject('stale-document')
    return live
  }

  return createWorkspaceDocumentAdapter({
    rollback: deps.rollback,
    subscribeChanges(handler) {
      changes.add(handler)
      const off = deps.subscribeChanges?.(handler)
      return () => {
        changes.delete(handler)
        off?.()
      }
    },
    getState: state,
    masking: 'applied',
    toolAccess: TOOL_ACCESS,
    beforeFirstWrite: deps.beforeFirstWrite,
    read(captured, request, signal) {
      const live = liveEditor()
      assertReady(live, captured.revision, captured.instanceId, signal)
      const frontmatter = safeFrontmatter()
      if (!live.markdown) reject('unsupported-read')
      let json: JSONContent = live.getJSON()
      if (request.scope === 'selection') {
        const frozen = frozenFor(captured)
        if (!frozen.valid) reject('stale-selection')
        const selection = frozen.selection
        if (selection.empty) reject('stale-selection')
        json = {
          type: 'doc',
          content: live.state.doc.slice(selection.from, selection.to).content.toJSON(),
        }
      }
      const body = live.markdown.serialize(redactJson(json))
      const text =
        request.scope === 'document' && frontmatter ? `---\n${frontmatter}\n---\n\n${body}` : body
      return { text, truncated: false }
    },
    createSkill(captured, signal): AgentSkill {
      const live = liveEditor()
      const frozen = frozenFor(captured)
      safeFrontmatter()
      let currentRevision = captured.revision
      const preparedDocumentId = documentId
      const original = createMarkdownSkill(
        () => live,
        deps.frontmatter,
        () => false,
      )

      // Temporarily restore the captured PM selection. Never focus the view;
      // the user's visible caret is restored and mapped through any local edit.
      const withSelection = <T>(operation: () => T, checkSignal = signal): T => {
        assertReady(live, currentRevision, captured.instanceId, checkSignal)
        let visible = live.state.selection
        const visibleMarks = live.state.storedMarks
        const map = ({
          transaction,
          appendedTransactions = [],
        }: {
          transaction: Transaction
          appendedTransactions?: Transaction[]
        }) => {
          for (const item of [transaction, ...appendedTransactions]) {
            if (!item.docChanged) continue
            const previous = frozen.selection
            frozen.selection = mapSelection(previous, item)
            if (
              (!previous.empty && frozen.selection.empty) ||
              (previous.empty && item.mapping.mapResult(previous.head, 1).deletedAcross)
            ) {
              frozen.valid = false
            }
            visible = visible.map(item.doc, item.mapping)
          }
        }
        live.on('transaction', map)
        try {
          if (!visible.eq(frozen.selection)) {
            live.view.dispatch(
              live.state.tr.setSelection(frozen.selection).setMeta('addToHistory', false),
            )
          }
          assertReady(live, currentRevision, captured.instanceId, checkSignal)
          return operation()
        } finally {
          live.off('transaction', map)
          if (
            !live.isDestroyed &&
            deps.getEditor() === live &&
            deps.getDocumentId() === preparedDocumentId &&
            instanceId() === captured.instanceId
          ) {
            if (
              visible.$from.doc === live.state.doc &&
              (!live.state.selection.eq(visible) || live.state.storedMarks !== visibleMarks)
            ) {
              live.view.dispatch(
                live.state.tr
                  .setSelection(visible)
                  .setStoredMarks(visibleMarks)
                  .setMeta('addToHistory', false),
              )
            }
            syncRevision(live)
            currentRevision = String(revision)
            frozen.revision = currentRevision
            selections.delete(frozen.id)
            selections.set(frozen.id, frozen)
          }
        }
      }

      return {
        id: original.id,
        get systemPrompt() {
          return `${original.systemPrompt}\n\nWorkspace document mode: only the listed local document tools are available. Image insertion, media search/generation, long-form writer, mind maps, file creation and saving are unavailable. Selection-addressed operations use the captured selection, never a later caret.`
        },
        tools: original.tools
          .filter((tool) => Object.hasOwn(TOOL_ACCESS, tool.name))
          .map((tool) =>
            tool.name === 'apply_ops'
              ? {
                  ...tool,
                  description: `${tool.description}\nWorkspace mode does not support insertImage or image content.`,
                }
              : tool,
          ),
        buildContext: () => withSelection(() => original.buildContext?.() ?? ''),
        executeTool(call, operationSignal): ToolExecution | Promise<ToolExecution> {
          assertReady(live, currentRevision, captured.instanceId, operationSignal)
          if (call.name === 'apply_ops') {
            if (!live.isEditable) reject('read-only')
            const parsed = validateOps(call.input.ops)
            if ('ops' in parsed) {
              if (
                !frozen.valid &&
                parsed.ops.some(
                  (op) =>
                    ('target' in op && op.target === 'selection') ||
                    ('after' in op && op.after === 'selection'),
                )
              ) {
                reject('stale-selection')
              }
              for (const op of parsed.ops) {
                if (op.op === 'setFrontmatter' && frontmatterHasPrivateMarkup(op.yaml))
                  reject('unsafe-content')
                if (op.op === 'insertImage')
                  return {
                    output: 'Image insertion is unavailable in workspace document mode.',
                    summary: call.name,
                    isError: true,
                  }
                if (
                  (op.op === 'insertContent' || op.op === 'replaceBlocks') &&
                  live.markdown &&
                  containsImage(live.markdown.parse(op.markdown))
                ) {
                  return {
                    output: 'Image content is unavailable in workspace document mode.',
                    summary: call.name,
                    isError: true,
                  }
                }
              }
            }
            safeFrontmatter()
          }
          return withSelection(() => original.executeTool(call, operationSignal), operationSignal)
        },
        verifyResponse: original.verifyResponse?.bind(original),
      }
    },
    dispose() {
      editor?.off('transaction', changed)
      selections.clear()
    },
  })
}
