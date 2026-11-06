import type { Lang } from '@genoffice/i18n'
import { WorkspaceDocumentWriter } from './doc-writer'
import { aiLangDirectiveFor } from '../../../../../docs/src/renderer/i18n/ai-language'
import {
  AgentLoop,
  composeSkills,
  createMcpPermissionStore,
  sanitizeAgentPayload,
  READ_SKILL_FILE_TOOL,
  USE_SKILL_TOOL,
  type AgentImage,
  type AgentSkill,
  type AgentTransport,
  type McpPermissionAnswer,
  type McpPermissionStore,
  type ToolExecution,
  type AgentTokenUsage,
  type MindMapTreeNode,
} from '@genoffice/agent-core'
import type { ChatAttachment, ChatMessage, ToolActivity } from '@genoffice/project-store'
import { isSkillCardSavedMessage, type AiSendGate } from '@genoffice/ui'
import type {
  WorkspaceAiBoundRun,
  WorkspaceAiRunState,
  WorkspaceConversationOrigin,
  WorkspaceProjectBinding,
  WorkspaceBoundDocument,
  WorkspaceDocumentIdentity,
} from '../../../shared/workspace-ai-api'
import type {
  WorkspaceExecutionController,
  WorkspaceRollbackHandle,
  WorkspaceRunHandle,
  WorkspaceStartedRun,
} from '../../../main/workspace-ai/execution-controller'
import {
  WorkspaceConversations,
  workspaceConversationKey,
  type WorkspaceComposer,
  type WorkspaceConversationRun,
} from './conversations'
import { isProjectTool, toolRecordOf } from '../kb-ask-project'

type Awaitable<T> = T | Promise<T>
const sameBound = (a: WorkspaceBoundDocument, b: WorkspaceBoundDocument): boolean =>
  a.document.windowId === b.document.windowId &&
  a.document.tabId === b.document.tabId &&
  a.document.documentId === b.document.documentId &&
  a.document.kind === b.document.kind &&
  a.document.generation === b.document.generation &&
  a.revision === b.revision &&
  a.selection?.id === b.selection?.id &&
  a.selection?.revision === b.selection?.revision
/** Only guarded, main-owned document operations. The future IPC bridge must preserve these handles. */
export interface WorkspaceExecutionBridge {
  enqueue(
    request: Parameters<WorkspaceExecutionController['enqueue']>[0],
    attachments?: ChatAttachment[],
  ): Awaitable<WorkspaceAiBoundRun>
  startNext(): Promise<WorkspaceStartedRun | null>
  updateWriter?: WorkspaceExecutionController['updateWriter']
  execute: WorkspaceExecutionController['execute']
  verifyResponse: WorkspaceExecutionController['verifyResponse']
  complete: WorkspaceExecutionController['complete']
  cancel(handle: WorkspaceRunHandle): Awaitable<void>
  state(runId: string): Awaitable<WorkspaceAiRunState | null>
  rollbackForRun(handle: WorkspaceRunHandle): Awaitable<WorkspaceRollbackHandle | null>
  rollback: WorkspaceExecutionController['rollback']
}
export interface WorkspaceRunPreparation {
  /** Already checked by the existing project trust/context service; must equal the queued binding. */
  project: WorkspaceProjectBinding | null
  skills: AgentSkill[]
  images?: AgentImage[]
  /** Model-only skill instructions; the person's text alone is shown and stored. */
  userText?: string
  /** Wrap the complete composition, including document tools, in the existing outbound guards. */
  guard?(skill: AgentSkill): AgentSkill
  beforeTurn?(): Promise<void>
  onStart?(): void
  dispose(): void
}
export interface WorkspaceRunHostDeps {
  lang?(): Lang
  execution: WorkspaceExecutionBridge
  load(
    origin: WorkspaceConversationOrigin,
  ): Promise<{ readonly messages: readonly ChatMessage[]; readonly foreign: boolean }>
  append(
    origin: WorkspaceConversationOrigin,
    message: Omit<ChatMessage, 'seq' | 'ts'>,
  ): Promise<'recorded' | 'saved' | 'pending' | 'memory' | 'unwritable' | void>
  /** The panel supplies useAiSendGate.admit/admitQueued; blocked sends remain in the composer. */
  admit(request: WorkspacePendingSend, queued: boolean): boolean
  gate(): AiSendGate
  /** Called only after the send gate admits the captured request, never on selection/open. */
  captureProject?(request: WorkspacePendingSend): Promise<WorkspaceProjectBinding | null>
  discardProject?(conversationId: string, project: WorkspaceProjectBinding): Promise<void>
  /** One existing createIpcTransport per conversation, with its original errors/context/usage wiring. */
  createTransport(id: string, purpose?: 'writer'): AgentTransport & { dispose?(): void }
  /** Existing attachment, project, skill, drive/outbound and MCP guards stay in this factory. */
  prepare(
    input: WorkspaceComposer,
    run: WorkspaceStartedRun,
    permissions: McpPermissionStore,
  ): Promise<WorkspaceRunPreparation>
  toolsAllowed?(): boolean
  onUsage?(conversationId: string, usage: AgentTokenUsage): void
  errorText(error: unknown): string
}
export interface WorkspacePendingSend {
  /** Human-readable echo of an editor batch; model instructions remain in input.text. */
  displayText?: string
  conversationId: string
  targetId: string
  referenceIds: readonly string[]
  input: WorkspaceComposer
}
interface Runtime {
  writer: WorkspaceDocumentWriter
  loop: AgentLoop
  permissions: McpPermissionStore
  active: Work | null
  cancel(): void
  dispose(): void
}
export interface WorkspaceRollbackCard {
  id: string
  label: string
  state: 'applied' | 'restoring' | 'reverted' | 'refused'
}
export interface WorkspaceMindMapCard {
  readonly id: string
  readonly document: WorkspaceDocumentIdentity
  readonly tree: MindMapTreeNode
}
interface RollbackCard extends WorkspaceRollbackCard {
  handle: WorkspaceRollbackHandle
}
interface Work {
  request: WorkspacePendingSend
  ticket: WorkspaceConversationRun
  input: WorkspaceComposer
  binding: WorkspaceAiBoundRun
  started: WorkspaceStartedRun | null
  preparation: WorkspaceRunPreparation | null
  skill: AgentSkill | null
  tools: ToolActivity[]
  finishing: boolean
  finished: boolean
  loopSettled: Promise<void> | null
  settleLoop(): void
  off(): void
  done: Promise<void>
  resolve(): void
}
export interface WorkspaceSubmittedRun {
  run: WorkspaceConversationRun
  done: Promise<void>
}

/** Runs only explicitly submitted work. Selection/opening and card mounting have no model side effects. */
export class WorkspaceRunHost {
  readonly conversations: WorkspaceConversations<Runtime>
  private readonly queue: Work[] = []
  private readonly cards = new Map<string, RollbackCard[]>()
  private readonly maps = new Map<string, WorkspaceMindMapCard[]>()
  private readonly listeners = new Set<() => void>()
  private revision = 0
  private pumping = false
  private restoring = false
  private closed = false
  private active: Work | null = null
  private readonly dirtyHistory = new Map<string, number>()
  private readonly historyReads = new Map<string, Promise<void>>()
  private readonly blockedHistory = new Set<string>()
  private readonly admissions = new Map<
    string,
    { run: WorkspaceConversationRun; cancelled: boolean }
  >()
  private admitting: Promise<unknown> = Promise.resolve()

  constructor(private readonly deps: WorkspaceRunHostDeps) {
    this.conversations = new WorkspaceConversations({
      load: deps.load,
      changed: () => this.notify(),
      createRuntime: (id, messages) => this.createRuntime(id, messages),
    })
  }
  subscribe = (listener: () => void): (() => void) => {
    this.listeners.add(listener)
    return () => {
      this.listeners.delete(listener)
    }
  }
  version = (): number => this.revision
  /** UI-only, bounded to this window session. Never enters model history or saved messages. */
  writing(id: string) {
    return this.conversations.runtime(id).writer.state
  }
  decideWriting(id: string, writerId: string, keep: boolean) {
    this.conversations.runtime(id).writer.answer(writerId, keep)
  }
  stopWriting(id: string, writerId: string) {
    this.conversations.runtime(id).writer.stopWriting(writerId)
  }
  mindMaps(id: string): readonly WorkspaceMindMapCard[] {
    return this.maps.get(id) ?? []
  }
  /** An explicit card save records only its existing saved-card marker in the original history. */
  async appendNote(id: string, message: Omit<ChatMessage, 'seq' | 'ts'>): Promise<void> {
    if (this.closed) return
    try {
      if ((await this.deps.append(this.conversations.state(id).origin, message)) === 'unwritable')
        throw new Error('history-unwritable')
      if (!this.closed) this.historyChanged(this.conversations.state(id).origin)
    } catch (error) {
      if (!this.closed) this.conversations.historyError(id, this.deps.errorText(error))
    }
  }
  rebindHistory(
    previous: WorkspaceConversationOrigin,
    origin: Extract<WorkspaceConversationOrigin, { kind: 'legacy-document' }>,
  ): void {
    if (this.closed) return
    for (const state of this.conversations.list()) {
      if (workspaceConversationKey(state.origin) !== workspaceConversationKey(previous)) continue
      try {
        this.conversations.rebindLegacy(state.id, origin)
      } catch (error) {
        this.blockedHistory.add(state.id)
        this.conversations.historyError(state.id, this.deps.errorText(error))
        for (const { run } of state.runs) void this.cancel(run)
      }
    }
    this.historyChanged(origin)
  }
  /** External records refresh only their original idle conversation, preserving its draft and loop. */
  historyChanged(origin: WorkspaceConversationOrigin): void {
    if (this.closed) return
    for (const state of this.conversations.list()) {
      if (workspaceConversationKey(state.origin) !== workspaceConversationKey(origin)) continue
      this.dirtyHistory.set(state.id, (this.dirtyHistory.get(state.id) ?? 0) + 1)
      void this.refreshHistory(state.id).catch(() => {})
    }
  }
  private async refreshHistory(id: string, queued = false): Promise<void> {
    const reading = this.historyReads.get(id)
    if (reading) {
      await reading
      return
    }
    const version = this.dirtyHistory.get(id)
    if (!version || this.closed) return
    const before = this.conversations.state(id)
    if (before.runs.some((r) => r.status === 'running') || (!queued && before.runs.length)) return
    const revision = this.conversations.historyRevision(id)
    let received = false
    const task = (async () => {
      const history = await this.deps.load(before.origin)
      received = true
      if (this.closed) return
      const now = this.conversations.state(id)
      if (
        this.dirtyHistory.get(id) !== version ||
        this.conversations.historyRevision(id) !== revision ||
        workspaceConversationKey(now.origin) !== workspaceConversationKey(before.origin) ||
        now.runs.some((r) => r.status === 'running')
      )
        return
      this.conversations
        .runtime(id)
        .loop.restore(
          history.messages
            .filter((m) => !isSkillCardSavedMessage(m))
            .map(({ role, text }) => ({ role, text })),
        )
      this.conversations.replaceHistory(id, history.messages, history.foreign)
      this.dirtyHistory.delete(id)
    })()
    this.historyReads.set(id, task)
    try {
      await task
    } finally {
      this.historyReads.delete(id)
      if (
        received &&
        !this.closed &&
        this.dirtyHistory.has(id) &&
        !this.conversations.state(id).runs.length
      )
        queueMicrotask(() => {
          void this.refreshHistory(id).catch(() => {})
        })
    }
  }
  busy = (): boolean => (this.active !== null && !this.active.finished) || this.restoring
  private notify(): void {
    this.revision++
    for (const listener of [...this.listeners]) listener()
  }

  private createRuntime(id: string, messages: readonly ChatMessage[]): Runtime {
    const transport = this.deps.createTransport(id)
    const permissions = createMcpPermissionStore()
    const off = permissions.subscribe(() => this.notify())
    const runtime: Runtime = {
      writer: new WorkspaceDocumentWriter(() => this.notify()),
      active: null,
      permissions,
      cancel: () => {
        permissions.denyAll()
        runtime.loop.cancel()
      },
      dispose: () => {
        off()
        permissions.denyAll()
        transport.dispose?.()
      },
      loop: new AgentLoop({
        transport: {
          stream: (request, callbacks) => {
            const work = runtime.active
            if (!work?.preparation?.beforeTurn) return transport.stream(request, callbacks)
            let cancelled = false
            let inner: ReturnType<AgentTransport['stream']> | undefined
            void work.preparation
              .beforeTurn()
              .then(() => {
                if (cancelled) return
                if (this.closed || work.finishing || runtime.active !== work)
                  throw new Error('Workspace execution stopped')
                inner = transport.stream(request, callbacks)
              })
              .catch((error: unknown) => {
                if (!cancelled) callbacks.onError(this.deps.errorText(error))
              })
            return {
              cancel() {
                if (cancelled) return
                cancelled = true
                if (inner) inner.cancel()
                else queueMicrotask(() => callbacks.onDone())
              },
            }
          },
        },
        toolsAllowed: this.deps.toolsAllowed,
        skill: {
          id: 'workspace-host',
          get systemPrompt() {
            return runtime.active?.skill?.systemPrompt ?? ''
          },
          get tools() {
            return runtime.active?.skill?.tools ?? []
          },
          buildContext: () => runtime.active?.skill?.buildContext?.() ?? '',
          executeTool: (call, signal) => {
            if (!runtime.active?.skill) throw new Error('No active workspace execution')
            return runtime.active.skill.executeTool(call, signal)
          },
          verifyResponse: (text, calls) =>
            runtime.active?.skill?.verifyResponse?.(text, calls) ?? null,
        },
        events: {
          onUsage: (usage) => this.deps.onUsage?.(id, usage),
          onText: (text) => {
            if (runtime.active)
              this.conversations.deliver(runtime.active.ticket, { progress: text })
          },
          onToolExecuted: ({ call, execution }) => {
            const work = runtime.active
            if (!work || work.finishing) return
            if (
              !execution.isError &&
              execution.display?.kind === 'mindmap' &&
              execution.display.tree
            ) {
              const card = {
                id: `${work.ticket.generation}:${call.id}`,
                document: structuredClone(work.binding.target.document),
                tree: structuredClone(execution.display.tree),
              }
              this.maps.set(id, [...(this.maps.get(id) ?? []), card].slice(-16))
            }
            work.tools.push(
              isProjectTool(call.name)
                ? toolRecordOf(call, execution)
                : {
                    name: call.name,
                    summary: execution.summary,
                    isError: execution.isError,
                    ...(call.name === 'read_attachment' ||
                    call.name === 'drive_search' ||
                    call.name === 'drive_read' ||
                    call.name === USE_SKILL_TOOL ||
                    call.name === READ_SKILL_FILE_TOOL
                      ? {}
                      : { output: execution.output.slice(0, 16_000) }),
                  },
            )
            this.notify()
          },
          onDone: (result) => {
            runtime.active?.settleLoop()
            if (runtime.active)
              void this.end(
                runtime.active,
                result.cancelled ? new Error('cancelled') : null,
                result.text,
              )
          },
          onError: (error) => {
            runtime.active?.settleLoop()
            if (runtime.active) void this.end(runtime.active, error)
          },
        },
      }),
    }
    runtime.loop.restore(
      messages.filter((m) => !isSkillCardSavedMessage(m)).map(({ role, text }) => ({ role, text })),
    )
    return runtime
  }

  /** Capture the draft now, then serialize async admission so later sends cannot jump ahead. */
  submit(
    id: string,
    targetId: string,
    referenceIds: readonly string[],
  ): Promise<WorkspaceSubmittedRun | null> {
    return this.resume({
      conversationId: id,
      targetId,
      referenceIds: [...referenceIds],
      input: this.conversations.state(id).composer,
    })
  }

  /** useAiSendGate.start returns the original held draft here, even after another conversation was selected. */
  resume(request: WorkspacePendingSend): Promise<WorkspaceSubmittedRun | null> {
    if (this.closed) return Promise.reject(new Error('Workspace host is closed'))
    const { conversationId: id, targetId } = request
    if (this.blockedHistory.has(id)) return Promise.resolve(null)
    const input = structuredClone(request.input)
    if (!input.text.trim() || !input.target) return Promise.resolve(null)
    input.text = sanitizeAgentPayload(input.text)
    const references = [...request.referenceIds]
    const pending = {
      conversationId: id,
      targetId,
      referenceIds: references,
      input,
      ...(request.displayText ? { displayText: sanitizeAgentPayload(request.displayText) } : {}),
    }
    if (!this.deps.admit(structuredClone(pending), false) || this.deps.gate().kind !== 'run')
      return Promise.resolve(null)
    const ticket = this.conversations.enqueue(id, crypto.randomUUID(), input)
    const admission = { run: ticket, cancelled: false }
    this.admissions.set(ticket.runId, admission)
    const task = this.admitting.then(async () => {
      try {
        if (this.closed || admission.cancelled) return null
        if (this.deps.captureProject) {
          input.project = await this.deps.captureProject(structuredClone(pending))
          if (this.closed || admission.cancelled) return null
          this.conversations.bindProject(ticket, input.project)
        }
        const binding = await this.deps.execution.enqueue(
          {
            conversationId: id,
            runId: ticket.runId,
            targetId,
            referenceIds: references,
            project: input.project,
          },
          input.attachments,
        )
        if (
          !sameBound(input.target!, binding.target) ||
          input.references.length !== binding.references.length ||
          input.references.some(
            (ref, index) =>
              !sameBound(ref, binding.references[index]!) ||
              ref.snapshotId !== binding.references[index]!.snapshotId,
          )
        ) {
          await this.deps.execution.cancel(binding)
          throw new Error('Queued document selection does not match the draft')
        }
        if (this.closed || admission.cancelled) {
          await this.deps.execution.cancel(binding)
          return null
        }
        let resolve!: () => void
        const done = new Promise<void>((finish) => {
          resolve = finish
        })
        const work: Work = {
          request: pending,
          ticket,
          input,
          binding,
          started: null,
          preparation: null,
          skill: null,
          tools: [],
          finishing: false,
          finished: false,
          loopSettled: null,
          settleLoop: () => {},
          off: () => {},
          done,
          resolve,
        }
        this.admissions.delete(ticket.runId)
        this.queue.push(work)
        void this.pump()
        return { run: ticket, done }
      } catch (error) {
        if (!this.closed && !admission.cancelled) {
          this.conversations.fail(ticket, this.deps.errorText(error))
          this.conversations.finish(ticket, true)
        }
        return null
      } finally {
        this.admissions.delete(ticket.runId)
        if (input.project) await this.deps.discardProject?.(id, input.project).catch(() => {})
      }
    })
    this.admitting = task.catch(() => {})
    return task
  }

  private async pump(): Promise<void> {
    if (this.pumping || this.closed || this.restoring) return
    this.pumping = true
    try {
      while (!this.closed && !this.restoring && this.queue.length) {
        const work = this.queue.shift()!
        if (work.finishing) continue
        this.active = work
        try {
          await this.refreshHistory(work.ticket.conversationId, true)
          if (this.dirtyHistory.has(work.ticket.conversationId))
            throw new Error('Conversation history changed while preparing')
          const state = await this.deps.execution.state(work.binding.runId)
          if (state?.status !== 'queued')
            throw new Error(state?.status === 'stopped' ? state.reason : 'Run is no longer queued')
          const admitted = this.deps.admit(structuredClone(work.request), true)
          const gate = this.deps.gate()
          if (!admitted || gate.kind !== 'run') {
            if (gate.kind === 'checking' || gate.kind === 'notice') {
              this.queue.unshift(work)
              this.active = null
              break
            }
            throw new Error('Send gate blocked the queued run')
          }
          const started = await this.deps.execution.startNext()
          if (
            !started ||
            started.binding.bindingId !== work.binding.bindingId ||
            started.binding.runId !== work.binding.runId
          )
            throw new Error('Unexpected workspace execution')
          work.started = started
          const runtime = this.conversations.runtime(work.ticket.conversationId)
          runtime.active = work
          const stop = () => {
            void this.end(work, new Error('Workspace execution stopped'))
          }
          started.signal.addEventListener('abort', stop, { once: true })
          work.off = () => started.signal.removeEventListener('abort', stop)
          if (started.signal.aborted) throw new Error('Workspace execution stopped')
          work.preparation = await this.deps.prepare(
            structuredClone(work.input),
            started,
            runtime.permissions,
          )
          if (work.finishing || this.closed) {
            this.releasePreparation(work)
            throw new Error('Workspace execution stopped')
          }
          if (
            work.preparation.project?.id !== started.binding.project?.id ||
            work.preparation.project?.contextVersion !== started.binding.project?.contextVersion
          )
            throw new Error('Project context changed')
          if (
            !this.deps.admit(structuredClone(work.request), true) ||
            this.deps.gate().kind !== 'run'
          )
            throw new Error('Send gate blocked the prepared run')
          const document: AgentSkill = {
            id: 'workspace-document',
            systemPrompt: started.target.systemPrompt,
            tools: [...started.target.tools],
            buildContext: () =>
              [
                started.target.context,
                ...started.references.map(
                  (ref) =>
                    `Explicit read-only reference ${JSON.stringify({ document: ref.binding.document, scope: ref.snapshot.scope, revision: ref.snapshot.state.revision })}:\n${ref.snapshot.text}`,
                ),
              ].join('\n\n'),
            executeTool: async (call, signal): Promise<ToolExecution> => {
              if (signal?.aborted || started.signal.aborted) throw new Error('cancelled')
              const executing = this.deps.execution.execute(
                work.binding,
                work.binding.target.document,
                call,
              )
              const update = this.deps.execution.updateWriter
              if (call.name !== 'write_document' || !update) return (await executing).result
              const controller = new AbortController()
              const writerSignal = AbortSignal.any([
                started.signal,
                controller.signal,
                ...(signal ? [signal] : []),
              ])
              const writerTransport = this.deps.createTransport(
                work.ticket.conversationId,
                'writer',
              )
              const writing = runtime.writer.run({
                call,
                instruction: work.input.text,
                directive: aiLangDirectiveFor(this.deps.lang?.() ?? 'en'),
                transport: writerTransport,
                onUsage: (usage) => this.deps.onUsage?.(work.ticket.conversationId, usage),
                signal: writerSignal,
                beforeTurn: async () => {
                  await work.preparation?.beforeTurn?.()
                },
                update: (value) => update.call(this.deps.execution, work.binding, call.id, value),
              })
              try {
                const first = await Promise.race([
                  executing.then((value) => ({ kind: 'executed' as const, value })),
                  writing.then(() => ({ kind: 'written' as const })),
                ])
                if (first.kind === 'executed') return first.value.result
                return (await executing).result
              } catch (error) {
                try {
                  await this.deps.execution.cancel(work.binding)
                } catch {
                  // The editor or main invalidation may already have stopped this run.
                }
                throw error
              } finally {
                controller.abort()
                await writing.catch(() => {})
                writerTransport.dispose?.()
              }
            },
            verifyResponse: (text, calls) =>
              this.deps.execution.verifyResponse(work.binding, text, calls),
          }
          const composed = composeSkills('workspace', '', [document, ...work.preparation.skills])
          work.skill = work.preparation.guard?.(composed) ?? composed
          if (this.dirtyHistory.has(work.ticket.conversationId))
            throw new Error('Conversation history changed while preparing')
          if (!this.conversations.start(work.ticket))
            throw new Error('Conversation execution is not available')
          const messages = this.conversations.state(work.ticket.conversationId).messages
          this.conversations.deliver(work.ticket, {
            message: {
              seq: (messages.at(-1)?.seq ?? -1) + 1,
              ts: new Date().toISOString(),
              role: 'user',
              text: work.request.displayText ?? work.input.text,
              attachments: work.input.attachments,
              skills: work.input.skills,
            },
          })
          work.preparation.onStart?.()
          work.loopSettled = new Promise<void>((resolve) => {
            work.settleLoop = resolve
          })
          runtime.loop.run(work.preparation.userText ?? work.input.text, work.preparation.images)
        } catch (error) {
          await this.end(work, error)
        }
        await work.done
        this.active = null
      }
    } finally {
      this.pumping = false
      this.notify()
    }
  }

  private async end(work: Work, error: unknown, text = ''): Promise<void> {
    if (work.finishing) return work.done
    work.finishing = true
    work.off()
    const id = work.ticket.conversationId
    try {
      if (error) {
        try {
          await this.deps.execution.cancel(work.binding)
        } catch {
          /* The main invalidation may already have stopped it. */
        }
        if (!this.closed) {
          const runtime = this.conversations.runtime(id)
          const composer = this.conversations.state(id).composer
          if (
            !work.loopSettled &&
            !composer.text &&
            !composer.attachments.length &&
            !composer.skills.length
          )
            this.conversations.edit(id, {
              composer: {
                ...composer,
                text: work.request.displayText ?? work.input.text,
                attachments: work.input.attachments,
                skills: work.input.skills,
              },
            })
          if (runtime.active === work && work.loopSettled) {
            runtime.permissions.denyAll()
            runtime.loop.cancel()
            await work.loopSettled
          }
        }
        if (!this.closed) this.conversations.fail(work.ticket, this.deps.errorText(error))
      } else await this.deps.execution.complete(work.binding)
      if (this.closed) return
      const rollback = await this.deps.execution.rollbackForRun(work.binding)
      if (rollback) {
        const cards = this.cards.get(id) ?? []
        cards.push({
          id: rollback.rollbackId,
          label: (work.request.displayText ?? work.input.text).slice(0, 80),
          state: 'applied',
          handle: rollback,
        })
        this.cards.set(id, cards.slice(-16))
      }
      if (text) {
        const message: Omit<ChatMessage, 'seq' | 'ts'> = {
          role: 'assistant',
          text,
          tools: work.tools,
        }
        const messages = this.conversations.state(id).messages
        this.conversations.deliver(work.ticket, {
          message: {
            ...message,
            seq: (messages.at(-1)?.seq ?? -1) + 1,
            ts: new Date().toISOString(),
          },
        })
        if (!error) {
          // Resolve the current source after first-save rebind, never a captured file path.
          const first = await this.deps.append(this.conversations.state(id).origin, {
            role: 'user',
            text: work.request.displayText ?? work.input.text,
            attachments: work.input.attachments,
            skills: work.input.skills,
          })
          if (first === 'unwritable') throw new Error('history-unwritable')
          if (
            !this.closed &&
            (await this.deps.append(this.conversations.state(id).origin, message)) === 'unwritable'
          )
            throw new Error('history-unwritable')
        }
      }
    } catch (failure) {
      try {
        await this.deps.execution.cancel(work.binding)
      } catch {}
      if (!this.closed) this.conversations.fail(work.ticket, this.deps.errorText(failure))
    } finally {
      if (!this.closed) {
        const runtime = this.conversations.runtime(id)
        if (runtime.active === work) runtime.permissions.denyAll()
        this.conversations.finish(work.ticket, !!error)
        if (runtime.active === work) runtime.active = null
      }
      this.releasePreparation(work)
      work.finished = true
      work.resolve()
      this.notify()
      void this.refreshHistory(id).catch(() => {})
    }
  }

  private releasePreparation(work: Work): void {
    const preparation = work.preparation
    work.preparation = null
    try {
      preparation?.dispose()
    } catch {
      /* Cleanup does not replay a finished model request. */
    }
  }

  async cancel(run: WorkspaceConversationRun): Promise<void> {
    const admission = this.admissions.get(run.runId)
    if (
      admission &&
      admission.run.controllerId === run.controllerId &&
      admission.run.conversationId === run.conversationId &&
      admission.run.generation === run.generation
    ) {
      admission.cancelled = true
      this.conversations.fail(admission.run, this.deps.errorText(new Error('cancelled')))
      this.conversations.finish(admission.run, true)
    }
    const work = [this.active, ...this.queue].find(
      (item) =>
        item?.ticket.controllerId === run.controllerId &&
        item.ticket.conversationId === run.conversationId &&
        item.ticket.runId === run.runId &&
        item.ticket.generation === run.generation,
    )
    if (work) {
      if (work.finishing) {
        try {
          await this.deps.execution.cancel(work.binding)
        } catch {
          // A fully completed run has no pending first save to cancel.
        }
      }
      await this.end(work, new Error('cancelled'))
    }
  }
  /** Called by the panel when the shared gate changes; it can resume only previously submitted work. */
  wake(): void {
    void this.pump()
  }
  permissions(id: string): McpPermissionStore {
    return this.conversations.runtime(id).permissions
  }
  answerPermission(id: string, requestId: string, answer: McpPermissionAnswer): void {
    if (this.closed) return
    this.permissions(id).answer(requestId, answer)
  }
  rollbackCards(id: string): WorkspaceRollbackCard[] {
    return (this.cards.get(id) ?? []).map(({ handle: _, ...card }) => ({ ...card }))
  }
  async rollback(id: string, cardId: string): Promise<void> {
    const card = this.cards.get(id)?.find((item) => item.id === cardId)
    if (!card || card.state !== 'applied' || this.closed || this.busy()) return
    this.restoring = true
    card.state = 'restoring'
    this.notify()
    try {
      await this.deps.execution.rollback(card.handle)
      card.state = 'reverted'
    } catch {
      card.state = 'refused'
    }
    this.restoring = false
    this.notify()
    void this.pump()
  }
  dispose(): void {
    if (this.closed) return
    this.closed = true
    for (const admission of this.admissions.values()) admission.cancelled = true
    for (const work of [this.active, ...this.queue]) {
      if (!work) continue
      work.off()
      void (async () => {
        await this.deps.execution.cancel(work.binding)
      })().catch(() => {})
      this.releasePreparation(work)
      work.resolve()
    }
    this.queue.length = 0
    this.conversations.dispose()
    this.cards.clear()
    this.maps.clear()
    this.listeners.clear()
  }
}
