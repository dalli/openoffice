import { sanitizeAgentPayload } from '@genoffice/agent-core'
import type { ChatAttachment, ChatMessage, ChatSkillChip } from '@genoffice/project-store'
import type {
  WorkspaceBoundDocument,
  WorkspaceConversationOrigin,
  WorkspaceProjectBinding,
  WorkspaceReference,
} from '../../../shared/workspace-ai-api'
import { workspaceConversationKey } from '../../../shared/workspace-ai-api'
export { workspaceConversationKey } from '../../../shared/workspace-ai-api'

export interface WorkspaceComposer {
  text: string
  attachments: ChatAttachment[]
  skills: ChatSkillChip[]
  target: WorkspaceBoundDocument | null
  references: WorkspaceReference[]
  project: WorkspaceProjectBinding | null
}

export interface WorkspaceConversationRuntime {
  /** Cancels this conversation's transport/loop; selecting another conversation never calls it. */
  cancel(): void
  dispose?(): void
}

export interface WorkspaceConversationRun {
  readonly controllerId: string
  readonly conversationId: string
  readonly runId: string
  readonly generation: number
}

export interface WorkspacePermission {
  id: string
  run: WorkspaceConversationRun
  payload: unknown
}

export interface WorkspaceConversationState {
  id: string
  origin: WorkspaceConversationOrigin
  messages: ChatMessage[]
  foreign: boolean
  composer: WorkspaceComposer
  scrollTop: number
  /** Conversation-local MCP choices; never restored from another conversation's records. */
  mcp: { disabledServers: string[]; alwaysAllowedTools: string[] }
  runs: Array<{
    run: WorkspaceConversationRun
    status: 'queued' | 'running'
    input: WorkspaceComposer
  }>
  permissions: WorkspacePermission[]
  progress: string
  error: string | null
}

interface RunEntry {
  ticket: WorkspaceConversationRun
  abort: AbortController
  permissionIds: Set<string>
}
interface Session<R> {
  historyRevision: number
  state: WorkspaceConversationState
  runtime: R
  runs: Map<number, RunEntry>
  permissions: Map<string, (allowed: boolean) => void>
}

export interface WorkspaceConversationsDeps<R extends WorkspaceConversationRuntime> {
  load(origin: WorkspaceConversationOrigin): Promise<{
    readonly messages: readonly ChatMessage[]
    readonly foreign: boolean
    readonly origin?: WorkspaceConversationOrigin
  }>
  /** Create one loop/history owner per conversation. This factory must not start a model request. */
  createRuntime(id: string, messages: readonly ChatMessage[]): R
  changed?(): void
}

/** No IPC, model calls or implicit persistence. A future panel host supplies the guarded bridges. */
export class WorkspaceConversations<R extends WorkspaceConversationRuntime> {
  private readonly id = globalThis.crypto.randomUUID()
  private readonly sessions = new Map<string, Session<R>>()
  private readonly aliases = new Map<string, string>()
  private readonly opening = new Map<string, Promise<string>>()
  private selected: string | null = null
  private generation = 0
  private closed = false

  constructor(private readonly deps: WorkspaceConversationsDeps<R>) {}

  get selectedId(): string | null {
    return this.selected
  }

  /** Opening reads history once. It does not select, send, cancel or rewrite an existing session. */
  open(
    origin: WorkspaceConversationOrigin,
    initialTarget: WorkspaceBoundDocument | null = null,
  ): Promise<string> {
    if (this.closed) return Promise.reject(new Error('Conversation controller is closed'))
    const key = workspaceConversationKey(origin)
    const existing = this.aliases.get(key) ?? key
    if (this.sessions.has(existing)) return Promise.resolve(existing)
    const pending = this.opening.get(key)
    if (pending) return pending
    const savedOrigin = structuredClone(origin)
    const target = structuredClone(initialTarget)
    const task = Promise.resolve()
      .then(async () => {
        const history = await this.deps.load(savedOrigin)
        if (this.closed) throw new Error('Conversation controller is closed')
        const resolvedKey = workspaceConversationKey(history.origin ?? savedOrigin)
        const resolved = this.aliases.get(resolvedKey) ?? resolvedKey
        if (this.sessions.has(resolved)) throw new Error('Cannot merge open conversation histories')
        const messages = structuredClone([...history.messages])
        const runtime = this.deps.createRuntime(key, structuredClone(messages))
        this.sessions.set(key, {
          historyRevision: 0,
          runtime,
          runs: new Map(),
          permissions: new Map(),
          state: {
            id: key,
            origin: structuredClone(history.origin ?? savedOrigin),
            messages,
            foreign: history.foreign,
            composer: {
              text: '',
              attachments: [],
              skills: [],
              target,
              references: [],
              project: null,
            },
            scrollTop: 0,
            mcp: { disabledServers: [], alwaysAllowedTools: [] },
            runs: [],
            permissions: [],
            progress: '',
            error: null,
          },
        })
        this.aliases.set(workspaceConversationKey(history.origin ?? savedOrigin), key)
        this.deps.changed?.()
        return key
      })
      .finally(() => this.opening.delete(key))
    this.opening.set(key, task)
    return task
  }

  private session(id: string): Session<R> {
    const session = this.sessions.get(id)
    if (this.closed || !session) throw new Error('Conversation is not open')
    return session
  }

  select(id: string): void {
    this.session(id)
    this.selected = id
    this.deps.changed?.()
  }

  /** Apply the existing first-save/move service's result without replacing the conversation/loop. */
  rebindLegacy(id: string, ref: { projectId: string; chatId: string }): void {
    const session = this.session(id)
    if (session.state.origin.kind !== 'legacy-document')
      throw new Error('Not a legacy document conversation')
    const origin: WorkspaceConversationOrigin = { kind: 'legacy-document', ...ref }
    const key = workspaceConversationKey(origin)
    const existing = this.aliases.get(key) ?? key
    if ((this.sessions.has(existing) && existing !== id) || this.opening.has(key)) {
      throw new Error('Cannot merge open conversation histories')
    }
    session.state.origin = origin
    this.aliases.set(key, id)
    this.deps.changed?.()
  }

  state(id: string): WorkspaceConversationState {
    return structuredClone(this.session(id).state)
  }
  runtime(id: string): R {
    return this.session(id).runtime
  }
  historyRevision(id: string): number {
    return this.session(id).historyRevision
  }
  historyError(id: string, error: string): void {
    this.session(id).state.error = error
    this.deps.changed?.()
  }
  replaceHistory(id: string, messages: readonly ChatMessage[], foreign: boolean): void {
    const session = this.session(id)
    if (session.state.runs.some((run) => run.status === 'running'))
      throw new Error('Cannot replace running history')
    session.state.messages = structuredClone([...messages])
    session.state.foreign = foreign
    session.historyRevision++
    this.deps.changed?.()
  }
  list(): WorkspaceConversationState[] {
    return [...this.sessions.keys()].map((id) => this.state(id))
  }

  edit(
    id: string,
    changes: Partial<Pick<WorkspaceConversationState, 'composer' | 'scrollTop' | 'mcp'>>,
  ): void {
    Object.assign(this.session(id).state, structuredClone(changes))
    this.deps.changed?.()
  }

  /** Called only after the host's send gate accepts an explicit send; input is frozen at enqueue. */
  enqueue(id: string, runId: string, draft?: WorkspaceComposer): WorkspaceConversationRun {
    const session = this.session(id)
    if (!runId || session.state.runs.some(({ run }) => run.runId === runId)) {
      throw new Error('Invalid or duplicate conversation run')
    }
    const ticket = Object.freeze({
      controllerId: this.id,
      conversationId: id,
      runId,
      generation: ++this.generation,
    })
    session.runs.set(ticket.generation, {
      ticket,
      abort: new AbortController(),
      permissionIds: new Set(),
    })
    const input = structuredClone(draft ?? session.state.composer)
    input.text = sanitizeAgentPayload(input.text)
    session.state.runs.push({ run: ticket, status: 'queued', input })
    this.deps.changed?.()
    return ticket
  }

  private matching(run: WorkspaceConversationRun): Session<R> | null {
    if (run.controllerId !== this.id) return null
    const session = this.closed ? undefined : this.sessions.get(run.conversationId)
    const entry = session?.runs.get(run.generation)
    return entry?.ticket.runId === run.runId ? session! : null
  }

  signal(run: WorkspaceConversationRun): AbortSignal | null {
    return this.matching(run)?.runs.get(run.generation)?.abort.signal ?? null
  }

  /** Main must have prepared this FIFO request and settled the previous editor before this call. */
  start(run: WorkspaceConversationRun): boolean {
    const session = this.matching(run)
    if (
      !session ||
      [...this.sessions.values()].some((s) => s.state.runs.some((r) => r.status === 'running'))
    )
      return false
    session.state.runs.find((r) => r.run.generation === run.generation)!.status = 'running'
    session.state.error = null
    this.deps.changed?.()
    return true
  }

  /** Late callbacks always carry the original ticket, never the current selection. */
  deliver(
    run: WorkspaceConversationRun,
    event: { message: ChatMessage } | { progress: string } | { error: string },
  ): boolean {
    const session = this.matching(run)
    if (
      !session ||
      !session.state.runs.some((r) => r.run.generation === run.generation && r.status === 'running')
    )
      return false
    if ('message' in event) {
      const message = structuredClone(event.message)
      if (message.role === 'user') message.text = sanitizeAgentPayload(message.text)
      session.state.messages.push(message)
      session.historyRevision++
    } else if ('progress' in event) session.state.progress = event.progress
    else session.state.error = event.error
    this.deps.changed?.()
    return true
  }

  ask(run: WorkspaceConversationRun, id: string, payload: unknown): Promise<boolean> {
    const session = this.matching(run)
    if (
      !session ||
      !session.state.runs.some((r) => r.run.generation === run.generation && r.status === 'running')
    )
      return Promise.resolve(false)
    const key = `${run.generation}:${id}`
    const entry = session.runs.get(run.generation)!
    if (!id || entry.permissionIds.has(id)) return Promise.resolve(false)
    const copied = structuredClone(payload)
    entry.permissionIds.add(id)
    const answer = new Promise<boolean>((resolve) => session.permissions.set(key, resolve))
    session.state.permissions.push({ id, run, payload: copied })
    this.deps.changed?.()
    return answer
  }

  answer(run: WorkspaceConversationRun, id: string, allowed: boolean): boolean {
    const session = this.matching(run)
    const key = `${run.generation}:${id}`
    const resolve = session?.permissions.get(key)
    if (!session || !resolve) return false
    session.permissions.delete(key)
    session.state.permissions = session.state.permissions.filter(
      (p) => p.run.generation !== run.generation || p.id !== id,
    )
    resolve(allowed)
    this.deps.changed?.()
    return true
  }

  fail(run: WorkspaceConversationRun, error: string): boolean {
    const session = this.matching(run)
    if (!session) return false
    session.state.error = error
    this.deps.changed?.()
    return true
  }

  finish(run: WorkspaceConversationRun, cancelled = false): boolean {
    const session = this.matching(run)
    if (!session) return false
    const entry = session.runs.get(run.generation)!
    const running = session.state.runs.some(
      (r) => r.run.generation === run.generation && r.status === 'running',
    )
    for (const permission of [...session.state.permissions]) {
      if (permission.run.generation === run.generation) this.answer(run, permission.id, false)
    }
    session.runs.delete(run.generation)
    session.state.runs = session.state.runs.filter((r) => r.run.generation !== run.generation)
    if (running) session.state.progress = ''
    // Invalidate callbacks before abort/cancel can synchronously deliver more events.
    if (cancelled) {
      entry.abort.abort()
      if (running) session.runtime.cancel()
    }
    this.deps.changed?.()
    return true
  }

  dispose(): void {
    if (this.closed) return
    this.closed = true
    const sessions = [...this.sessions.values()]
    this.sessions.clear()
    this.aliases.clear()
    this.selected = null
    let failure: unknown
    for (const session of sessions) {
      for (const entry of session.runs.values()) entry.abort.abort()
      for (const resolve of session.permissions.values()) resolve(false)
      try {
        session.runtime.cancel()
      } catch (error) {
        failure ??= error
      }
      try {
        session.runtime.dispose?.()
      } catch (error) {
        failure ??= error
      }
    }
    if (failure) throw failure
  }
}
