import type { AgentSkill, ExecutedToolCall } from './skill'
import type { AgentToolCall, AgentToolDef, ToolExecution } from './types'
import { coerceArgumentFields, invalidArgumentFields, missingRequiredFields } from './loop'

export interface WorkspaceAdapterState {
  /** Changes on replacement/reload, including loading another file into the same editor. */
  readonly instanceId: string
  readonly revision: string
  /** Opaque editor-owned selection captured with this revision; never a live DOM selection. */
  readonly selectionId: string | null
  readonly readOnly: boolean
}

export interface WorkspaceReadRequest {
  readonly scope: 'document' | 'selection'
  readonly maxChars?: number
}

export interface WorkspaceReadContent {
  /** Editor-approved text; there is no raw-text or disk fallback here. */
  readonly text: string
  readonly truncated: boolean
}

export interface WorkspaceReadSnapshot extends WorkspaceReadContent {
  readonly state: WorkspaceAdapterState
  readonly scope: WorkspaceReadRequest['scope']
  readonly masking: 'applied' | 'unsupported'
}

export type WorkspaceAdapterErrorKind =
  | 'not-ready'
  | 'disposed'
  | 'busy'
  | 'stale-document'
  | 'stale-selection'
  | 'read-only'
  | 'unsupported-read'
  | 'unsafe-content'
  | 'cancelled'
  | 'unknown-tool'
  | 'invalid-call'
  | 'duplicate-call'
  | 'released'

export class WorkspaceAdapterError extends Error {
  constructor(readonly kind: WorkspaceAdapterErrorKind) {
    super(`Workspace document request rejected: ${kind}`)
    this.name = 'WorkspaceAdapterError'
  }
}

export interface WorkspacePrepareRequest {
  readonly runId: string
  readonly mode: 'edit' | 'reference'
  readonly state: WorkspaceAdapterState
}

export interface WorkspacePreparedDocument {
  readonly id: string
  readonly runId: string
  readonly state: WorkspaceAdapterState
  readonly systemPrompt: string
  readonly context: string
  readonly tools: readonly AgentToolDef[]
  readonly rollbackId?: string | null
  execute(
    call: AgentToolCall,
    signal?: AbortSignal,
  ): Promise<{ result: ToolExecution; state: WorkspaceAdapterState }>
  verifyResponse(text: string, executed: readonly ExecutedToolCall[]): Promise<string | null>
  release(): void
}

export interface WorkspaceDocumentAdapter {
  readonly toolAccess: Readonly<Record<string, 'read' | 'write'>>
  readonly masking: 'applied' | 'unsupported'
  readonly observesChanges: boolean
  subscribeChanges(handler: () => void): () => void
  capture(signal?: AbortSignal): Promise<WorkspaceAdapterState>
  read(
    state: WorkspaceAdapterState,
    request: WorkspaceReadRequest,
    signal?: AbortSignal,
  ): Promise<WorkspaceReadSnapshot>
  prepare(
    request: WorkspacePrepareRequest,
    signal?: AbortSignal,
  ): Promise<WorkspacePreparedDocument>
  rollback?(
    id: string,
    state: WorkspaceAdapterState,
    signal?: AbortSignal,
  ): Promise<WorkspaceAdapterState>
  dispose(): void
}

export interface WorkspaceDocumentAdapterDeps {
  /** Unsupported formats cannot claim that their snapshots apply masking. */
  readonly masking?: 'applied' | 'unsupported'
  getState(): WorkspaceAdapterState | Promise<WorkspaceAdapterState>
  read(
    state: WorkspaceAdapterState,
    request: WorkspaceReadRequest,
    signal?: AbortSignal,
  ): WorkspaceReadContent | Promise<WorkspaceReadContent>
  /** Must restore the captured selection into the skill, without changing the user's visible selection. */
  createSkill(state: WorkspaceAdapterState, signal?: AbortSignal): AgentSkill | Promise<AgentSkill>
  /** Explicit local document capabilities; unlisted tools never cross this boundary. */
  readonly toolAccess: Readonly<Record<string, 'read' | 'write'>>
  /** Metadata-only invalidation, including edits outside the document body. */
  subscribeChanges?(handler: () => void): () => void
  /** Original editor undo/change-tracking hook, once before the first write. */
  beforeFirstWrite?(signal?: AbortSignal): void | Promise<void>
  /** Editor-local immutable snapshot. Restore must synchronously validate/apply, with no async gap. */
  rollback?: { capture(): unknown; restore(snapshot: unknown): void }
  dispose?(): void
}

export const WORKSPACE_READ_MAX_CHARS = 200_000

const sameDocument = (a: WorkspaceAdapterState, b: WorkspaceAdapterState): boolean =>
  a.instanceId === b.instanceId && a.revision === b.revision
const fail = (kind: WorkspaceAdapterErrorKind): never => {
  throw new WorkspaceAdapterError(kind)
}
const checkAbort = (signal?: AbortSignal): void => {
  if (signal?.aborted) fail('cancelled')
}

/**
 * Renderer-only execution boundary. It never reads files, selects a provider or
 * calls a model. Async editor bridges must also validate the captured revision
 * at the actual mutation, after every await/permission prompt.
 */
export function createWorkspaceDocumentAdapter(
  deps: WorkspaceDocumentAdapterDeps,
): WorkspaceDocumentAdapter {
  const toolAccess = Object.freeze({ ...deps.toolAccess })
  const adapterId = crypto.randomUUID()
  const masking = deps.masking ?? 'unsupported'
  let disposed = false
  let active = false
  let sequence = 0
  let releaseActive: (() => void) | null = null
  const subscriptions = new Set<() => void>()
  const rollbacks = new Map<string, { snapshot: unknown; state: WorkspaceAdapterState }>()
  const live = (): void => {
    if (disposed) fail('disposed')
  }
  const capture = async (signal?: AbortSignal): Promise<WorkspaceAdapterState> => {
    live()
    checkAbort(signal)
    const state = await deps.getState()
    live()
    checkAbort(signal)
    if (
      typeof state.instanceId !== 'string' ||
      !state.instanceId ||
      typeof state.revision !== 'string' ||
      !state.revision ||
      typeof state.readOnly !== 'boolean' ||
      (state.selectionId !== null && (typeof state.selectionId !== 'string' || !state.selectionId))
    )
      fail('not-ready')
    return Object.freeze({ ...state })
  }
  const assertCurrent = async (state: WorkspaceAdapterState, signal?: AbortSignal) => {
    const now = await capture(signal)
    if (!sameDocument(state, now)) fail('stale-document')
    return now
  }
  return {
    toolAccess,
    masking,
    observesChanges: !!deps.subscribeChanges,
    subscribeChanges(handler) {
      live()
      const off = deps.subscribeChanges?.(handler) ?? (() => {})
      const unsubscribe = () => {
        if (!subscriptions.delete(unsubscribe)) return
        off()
      }
      subscriptions.add(unsubscribe)
      return unsubscribe
    },
    capture,
    async rollback(id, state, signal) {
      live()
      if (active) fail('busy')
      const saved = rollbacks.get(id)
      if (!saved || !deps.rollback || !sameDocument(saved.state, state)) fail('released')
      active = true
      try {
        const current = await assertCurrent(state, signal)
        if (current.readOnly) fail('read-only')
        live()
        checkAbort(signal)
        deps.rollback!.restore(saved!.snapshot)
        rollbacks.delete(id)
        return await capture(signal)
      } finally {
        active = false
      }
    },
    async read(state, request, signal) {
      if (request.scope !== 'document' && request.scope !== 'selection') fail('unsupported-read')
      if (request.scope === 'selection' && !state.selectionId) fail('stale-selection')
      const maximum = request.maxChars ?? WORKSPACE_READ_MAX_CHARS
      if (!Number.isSafeInteger(maximum) || maximum <= 0) fail('unsupported-read')
      const maxChars = Math.min(maximum, WORKSPACE_READ_MAX_CHARS)
      await assertCurrent(state, signal)
      const content = await deps.read(state, { ...request, maxChars }, signal)
      await assertCurrent(state, signal)
      if (typeof content.text !== 'string' || typeof content.truncated !== 'boolean')
        fail('unsafe-content')
      return {
        text: content.text.slice(0, maxChars),
        truncated: content.truncated || content.text.length > maxChars,
        state: Object.freeze({ ...state }),
        scope: request.scope,
        masking,
      }
    },
    async prepare(request, signal) {
      live()
      if (active) fail('busy')
      if (!request.runId || (request.mode !== 'edit' && request.mode !== 'reference'))
        fail('invalid-call')
      const runId = request.runId
      const mode = request.mode
      active = true
      const controller = new AbortController()
      const combined = signal ? AbortSignal.any([signal, controller.signal]) : controller.signal
      let released = false
      let preparing = true
      let executing = false
      let state = Object.freeze({ ...request.state })
      const release = () => {
        if (released) return
        released = true
        controller.abort()
        if (!executing && !preparing) active = false
        if (releaseActive === release) releaseActive = null
      }
      releaseActive = release
      try {
        await assertCurrent(state, combined)
        const skill = await deps.createSkill(state, combined)
        live()
        checkAbort(combined)
        const context = skill.buildContext?.() ?? ''
        const systemPrompt = skill.systemPrompt
        await assertCurrent(state, combined)
        const allowed = (name: string): boolean =>
          Object.hasOwn(toolAccess, name) &&
          (toolAccess[name] === 'read' ||
            (mode === 'edit' && !state.readOnly && toolAccess[name] === 'write'))
        const tools = structuredClone(skill.tools.filter((tool) => allowed(tool.name)))
        if (new Set(tools.map((tool) => tool.name)).size !== tools.length) fail('invalid-call')
        let undoCaptured = false
        let rollbackSnapshot: unknown
        let rollbackId: string | null = null
        const calls = new Set<string>()
        const ready = (): void => {
          live()
          if (released) fail('released')
          checkAbort(combined)
        }
        ready()
        return {
          id: `workspace-document-${adapterId}-${++sequence}`,
          runId,
          get rollbackId() {
            return rollbackId
          },
          get state() {
            return Object.freeze({ ...state })
          },
          systemPrompt,
          context,
          get tools() {
            return structuredClone(tools)
          },
          async execute(raw, callerSignal) {
            ready()
            if (executing) fail('busy')
            if (!raw.id || calls.has(raw.id)) fail(raw.id ? 'duplicate-call' : 'invalid-call')
            const tool = tools.find((t) => t.name === raw.name)
            if (!tool || !allowed(raw.name) || !skill.tools.some((t) => t.name === raw.name))
              fail('unknown-tool')
            if (
              raw.inputError ||
              raw.truncated ||
              !raw.input ||
              typeof raw.input !== 'object' ||
              Array.isArray(raw.input)
            )
              fail('invalid-call')
            const call = { ...raw, input: coerceArgumentFields(tool, structuredClone(raw.input)) }
            if (
              missingRequiredFields(tool, call.input).length ||
              invalidArgumentFields(tool, call.input).length
            )
              fail('invalid-call')
            const operationSignal = callerSignal
              ? AbortSignal.any([combined, callerSignal])
              : combined
            executing = true
            calls.add(call.id)
            try {
              const before = await assertCurrent(state, operationSignal)
              if (toolAccess[call.name] === 'write') {
                if (before.readOnly || mode === 'reference') fail('read-only')
                if (!undoCaptured) {
                  await deps.beforeFirstWrite?.(operationSignal)
                  undoCaptured = true
                  if ((await assertCurrent(state, operationSignal)).readOnly) fail('read-only')
                  rollbackSnapshot = deps.rollback?.capture()
                }
              }
              ready()
              checkAbort(operationSignal)
              const result = await skill.executeTool(call, operationSignal)
              const after = await capture(operationSignal)
              ready()
              if (
                after.instanceId !== state.instanceId ||
                (toolAccess[call.name] === 'read' &&
                  (result.mutated || !sameDocument(state, after))) ||
                (!result.mutated && !sameDocument(state, after)) ||
                (result.mutated && sameDocument(state, after))
              )
                fail('stale-document')
              state = Object.freeze({ ...after, selectionId: state.selectionId })
              if (result.mutated && deps.rollback) {
                rollbackId ??= crypto.randomUUID()
                rollbacks.set(rollbackId, { snapshot: rollbackSnapshot, state })
                while (rollbacks.size > 16) rollbacks.delete(rollbacks.keys().next().value!)
              }
              return { result, state }
            } catch (error) {
              release()
              throw error
            } finally {
              executing = false
              if (released) active = false
            }
          },
          async verifyResponse(text, executed) {
            ready()
            if (executing) fail('busy')
            executing = true
            try {
              await assertCurrent(state, combined)
              ready()
              const result = (await skill.verifyResponse?.(text, executed)) ?? null
              await assertCurrent(state, combined)
              ready()
              return result
            } finally {
              executing = false
              if (released) active = false
            }
          },
          release,
        }
      } catch (error) {
        release()
        throw error
      } finally {
        preparing = false
        if (released) active = false
      }
    },
    dispose() {
      if (disposed) return
      disposed = true
      releaseActive?.()
      rollbacks.clear()
      for (const unsubscribe of [...subscriptions]) unsubscribe()
      deps.dispose?.()
    },
  }
}
