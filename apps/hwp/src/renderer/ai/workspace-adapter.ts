import {
  createWorkspaceDocumentAdapter,
  WorkspaceAdapterError,
  type AgentSkill,
  type AgentToolCall,
  type WorkspaceAdapterState,
  type WorkspaceDocumentAdapter,
} from '@genoffice/agent-core'
import type {
  AgentSegment,
  AgentWorkspaceState,
  AgentWorkspaceVersion,
} from '../../../studio-plugin/agent-plugin.mjs'
import type { AgentBridge } from './agent-bridge'
import { createHwpSkill, type HwpSelection, type HwpSkillDeps } from './hwp-skill'

/** No files, network, images, provider calls or save operations enter this adapter. */
export const HWP_WORKSPACE_TOOL_ACCESS = {
  read_document: 'read',
  read_memos: 'read',
  find_text: 'read',
  read_selection: 'read',
  read_style: 'read',
  list_tables: 'read',
  read_table: 'read',
  list_fields: 'read',
  list_notes: 'read',
  read_note_settings: 'read',
  read_page_setup: 'read',
  read_formatting: 'read',
  replace_text: 'write',
  replace_all: 'write',
  insert_paragraphs: 'write',
  delete_paragraph: 'write',
  insert_table: 'write',
  format_text: 'write',
  replace_font: 'write',
  format_paragraph: 'write',
  replace_selection: 'write',
  format_selection: 'write',
  create_style: 'write',
  edit_style: 'write',
  merge_cells: 'write',
  split_cell: 'write',
  format_cells: 'write',
  fill_cells: 'write',
  check_boxes: 'write',
  add_to_table: 'write',
  remove_from_table: 'write',
  delete_table: 'write',
  fill_field: 'write',
  insert_toc: 'write',
  refresh_toc: 'write',
  insert_note: 'write',
  edit_note: 'write',
  delete_note: 'write',
  set_note_settings: 'write',
  set_page_setup: 'write',
  format_look: 'write',
} as const satisfies Readonly<Record<string, 'read' | 'write'>>

export interface HwpWorkspaceAdapterDeps {
  bridge: AgentBridge
  label: HwpSkillDeps['label']
  onChange: HwpSkillDeps['onChange']
  confirmMemoRemoval(
    refusal: Parameters<HwpSkillDeps['confirmMemoRemoval']>[0],
    signal?: AbortSignal,
  ): ReturnType<HwpSkillDeps['confirmMemoRemoval']>
  readOnly?(): boolean
  /** Retain the existing foreign-change baseline for the panel's guarded undo action. */
  onUndoBaseline?(foreignChanges: number): void
  subscribeChanges?(handler: () => void): () => void
}

interface CapturedSelection {
  instanceId: string
  revision: string
  scope: HwpSelection
}

const abort = (signal?: AbortSignal): void => {
  if (signal?.aborted) throw new WorkspaceAdapterError('cancelled')
}
const READ_METHODS = new Set([
  'foreignChanges',
  'readText',
  'findText',
  'listTables',
  'readTable',
  'listFields',
  'listNotes',
  'noteSettings',
  'pageSetup',
  'formatting',
  'listStyles',
  'readStyle',
  'listMemos',
  'selection',
  'readSelectionSnapshot',
])
let adapterSequence = 0

/**
 * A renderer-only adapter around the existing skill. Plugin version checks happen
 * in the same task as each document operation, including the retry after memo approval.
 * HWP has no masking-marker feature; reads explicitly report unsupported masking.
 */
export function createHwpWorkspaceAdapter(deps: HwpWorkspaceAdapterDeps): WorkspaceDocumentAdapter {
  const adapterId = ++adapterSequence
  const selections = new Map<string, CapturedSelection>()
  let sequence = 0
  let lastSelectionKey = ''
  let lastSelectionId: string | null = null
  let undoBaseline = 0
  const rollbacks = new Set<string>()
  let disposed = false

  const releaseRollback = (id: string): void => {
    rollbacks.delete(id)
    void deps.bridge.workspaceReleaseRollback(id).catch(() => {})
  }

  const readNative = async (signal?: AbortSignal): Promise<AgentWorkspaceState> => {
    abort(signal)
    const native = await deps.bridge.workspaceState()
    abort(signal)
    if (!native.instance || !native.revision || native.masking !== 'unsupported') {
      throw new WorkspaceAdapterError('unsafe-content')
    }
    return native
  }

  const selectionFor = (state: WorkspaceAdapterState): CapturedSelection | null => {
    if (!state.selectionId) return null
    const captured = selections.get(state.selectionId)
    if (
      !captured ||
      captured.instanceId !== state.instanceId ||
      captured.revision !== state.revision
    ) {
      throw new WorkspaceAdapterError('stale-selection')
    }
    return captured
  }

  const versionOf = (state: WorkspaceAdapterState): AgentWorkspaceVersion => ({
    instance: state.instanceId,
    revision: state.revision,
  })

  const invoke = async (
    expected: AgentWorkspaceVersion,
    method: string,
    args: unknown[],
    signal?: AbortSignal,
  ) => {
    abort(signal)
    const reply = await deps.bridge.workspaceInvoke(expected, method, args)
    abort(signal)
    if (!reply.ok || reply.instance !== expected.instance) {
      throw new WorkspaceAdapterError(
        reply.ok || reply.reason === 'changed' ? 'stale-document' : 'unknown-tool',
      )
    }
    return reply
  }

  return createWorkspaceDocumentAdapter({
    masking: 'unsupported',
    subscribeChanges: deps.subscribeChanges,
    toolAccess: HWP_WORKSPACE_TOOL_ACCESS,
    async getState() {
      const native = await readNative()
      if (!native.selection.ok && native.selection.reason !== 'none') {
        throw new WorkspaceAdapterError('unsupported-read')
      }
      const selected = native.selection.ok && !native.selection.own ? native.selection : null
      const key = selected ? JSON.stringify([native.instance, native.revision, selected]) : ''
      if (key !== lastSelectionKey || (key && !selections.has(lastSelectionId!))) {
        lastSelectionKey = key
        lastSelectionId = selected ? `hwp-selection-${adapterId}-${++sequence}` : null
        if (selected) {
          selections.set(lastSelectionId!, {
            instanceId: native.instance,
            revision: native.revision,
            scope: structuredClone({
              segments: selected.segments,
              inTable: selected.inTable,
              contiguous: selected.contiguous,
            }),
          })
          // Captures are ephemeral handles, not another copy of the document history.
          while (selections.size > 64) selections.delete(selections.keys().next().value!)
        }
      }
      return {
        instanceId: native.instance,
        revision: native.revision,
        selectionId: lastSelectionId,
        readOnly: native.readOnly === true || deps.readOnly?.() === true,
      }
    },
    async refreshSelection(state, current, signal) {
      const captured = state.selectionId ? selections.get(state.selectionId) : undefined
      if (
        !captured ||
        captured.instanceId !== current.instanceId ||
        captured.revision !== current.revision
      )
        throw new WorkspaceAdapterError('stale-selection')
      await invoke(versionOf(current), 'readSelectionSnapshot', [captured.scope.segments], signal)
      return state.selectionId!
    },
    async read(state, request, signal) {
      const selection = selectionFor(state)
      if (request.scope === 'selection' && !selection)
        throw new WorkspaceAdapterError('stale-selection')
      const reply = await invoke(
        versionOf(state),
        request.scope === 'selection' ? 'readSelectionSnapshot' : 'readText',
        request.scope === 'selection' ? [selection!.scope.segments] : [],
        signal,
      )
      if (typeof reply.value !== 'string') throw new WorkspaceAdapterError('unsafe-content')
      return { text: reply.value, truncated: false }
    },
    async createSkill(state, signal): Promise<AgentSkill> {
      const captured = selectionFor(state)
      let expected = versionOf(state)
      let operationSignal = signal
      // This bridge is private to one prepared session; a different session cannot retarget it.
      const guardedBridge = new Proxy({} as AgentBridge, {
        get(_target, method: string) {
          return async (...args: unknown[]) => {
            if (deps.readOnly?.() === true && !READ_METHODS.has(method)) {
              throw new WorkspaceAdapterError('read-only')
            }
            const reply = await invoke(expected, method, args, operationSignal)
            expected = { instance: reply.instance, revision: reply.revision }
            undoBaseline = reply.foreignChanges
            if (captured) {
              captured.revision = reply.revision
              if (method === 'replaceSelection') {
                const value = reply.value as { ok?: boolean; segments?: AgentSegment[] }
                if (value?.ok === true && Array.isArray(value.segments)) {
                  captured.scope.segments = structuredClone(value.segments)
                }
              }
            }
            return reply.value
          }
        },
      })
      const hwp = createHwpSkill({
        bridge: guardedBridge,
        label: deps.label,
        onChange: deps.onChange,
        confirmMemoRemoval: async (refusal) => {
          abort(operationSignal)
          const approved = await deps.confirmMemoRemoval(refusal, operationSignal)
          abort(operationSignal)
          // The retry's guarded bridge still carries the pre-question document revision.
          return approved
        },
        imageGenAvailable: () => false,
        pictureAnalysisAvailable: () => false,
      })
      hwp.setSelection(captured ? structuredClone(captured.scope) : null)
      await hwp.refreshContext()
      abort(signal)
      return {
        ...hwp.skill,
        // Accessors stay live, preserving the original skill's capability checks.
        get tools() {
          return hwp.skill.tools.filter((tool) =>
            Object.hasOwn(HWP_WORKSPACE_TOOL_ACCESS, tool.name),
          )
        },
        get systemPrompt() {
          return (
            hwp.skill.systemPrompt +
            '\n\nThis workspace adapter provides local document tools only. Network, images, attachments, file creation and saving are unavailable. Do not call or promise them.'
          )
        },
        async executeTool(call: AgentToolCall, callSignal?: AbortSignal) {
          operationSignal =
            signal && callSignal ? AbortSignal.any([signal, callSignal]) : (callSignal ?? signal)
          try {
            abort(operationSignal)
            // read_selection in the original skill reads its frozen cache: verify it still stands.
            if (call.name === 'read_selection' && captured) {
              await invoke(
                expected,
                'readSelectionSnapshot',
                [captured.scope.segments],
                operationSignal,
              )
            }
            return await hwp.skill.executeTool(call, operationSignal)
          } finally {
            operationSignal = signal
          }
        },
      }
    },
    beforeFirstWrite(signal) {
      abort(signal)
      deps.onUndoBaseline?.(undoBaseline)
    },
    rollback: {
      async capture(state, signal) {
        abort(signal)
        const reply = await deps.bridge.workspaceCaptureRollback(versionOf(state))
        if (!reply.ok) {
          throw new WorkspaceAdapterError(
            reply.reason === 'changed' ? 'stale-document' : 'not-ready',
          )
        }
        if (disposed || signal?.aborted) {
          releaseRollback(reply.id)
          throw new WorkspaceAdapterError(disposed ? 'released' : 'cancelled')
        }
        rollbacks.add(reply.id)
        while (rollbacks.size > 16) releaseRollback(rollbacks.values().next().value!)
        return {
          id: reply.id,
          selections: [...selections]
            .filter(([, selection]) => selection.revision === state.revision)
            .map(([id, selection]) => [id, structuredClone(selection)]),
        }
      },
      async restore(snapshot, state, signal) {
        abort(signal)
        if (deps.readOnly?.() === true) throw new WorkspaceAdapterError('read-only')
        const saved = snapshot as { id: string; selections: Array<[string, CapturedSelection]> }
        if (!saved || typeof saved.id !== 'string' || !rollbacks.has(saved.id)) {
          throw new WorkspaceAdapterError('released')
        }
        const reply = await deps.bridge.workspaceRollback(versionOf(state), saved.id)
        rollbacks.delete(saved.id)
        if (!reply.ok) {
          throw new WorkspaceAdapterError(
            reply.reason === 'changed' ? 'stale-document' : 'released',
          )
        }
        if (reply.instance !== state.instanceId) throw new WorkspaceAdapterError('stale-document')
        for (const [id, selection] of saved.selections) {
          if (selection.instanceId === reply.instance)
            selections.set(id, { ...selection, revision: reply.revision })
        }
        while (selections.size > 64) selections.delete(selections.keys().next().value!)
      },
    },
    dispose() {
      disposed = true
      for (const id of rollbacks) releaseRollback(id)
      selections.clear()
      lastSelectionKey = ''
      lastSelectionId = null
    },
  })
}
