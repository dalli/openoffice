import {
  createIpcTransport,
  type AgentTokenUsage,
  type IpcStreamError,
  type IpcSubscriptionUsage,
  type IpcStreamChunk,
} from '@genoffice/agent-core'
import {
  localContextLimit,
  localModelCapabilities,
  type AiSettings,
} from '@genoffice/ai-provider/browser'
import type { Lang } from '@genoffice/i18n'
import {
  addTokenUsage,
  aiContextConfirmText,
  aiRunErrorText,
  type AiTokenUsageView,
} from '@genoffice/ui'
import type { WorkspacePanelClient } from './panel-client'
import type { WorkspaceRunPreparation } from './run-host'
import type { WorkspaceStartedRun } from '../../../main/workspace-ai/execution-controller'
import { mergeSubscriptionUsage } from '../../../../../docs/src/renderer/ai/run-state'

export interface WorkspaceConnectionState {
  error: IpcStreamError | null
  subscriptionUsage: IpcSubscriptionUsage[]
  truncated: boolean
  usage: AiTokenUsageView | null
}
const empty = (): WorkspaceConnectionState => ({
  error: null,
  subscriptionUsage: [],
  truncated: false,
  usage: null,
})

/** Per-conversation transport feedback. Opening/selecting a conversation sends nothing. */
export class WorkspaceConnection {
  private readonly states = new Map<string, WorkspaceConnectionState>()
  private readonly runs = new Map<string, WorkspaceStartedRun>()
  private readonly listeners = new Set<() => void>()
  private revision = 0
  constructor(
    private readonly deps: {
      settings(): AiSettings | null
      lang(): Lang
      unknownError(): string
      confirm(text: string): boolean
      onError?(error: IpcStreamError): void
    },
  ) {}
  subscribe = (listener: () => void) => {
    this.listeners.add(listener)
    return () => {
      this.listeners.delete(listener)
    }
  }
  version = () => this.revision
  state(id: string): WorkspaceConnectionState {
    return this.states.get(id) ?? empty()
  }
  private patch(id: string, patch: Partial<WorkspaceConnectionState>) {
    this.states.set(id, { ...this.state(id), ...patch })
    this.revision++
    for (const listener of this.listeners) listener()
  }
  dismissTruncated(id: string) {
    this.patch(id, { truncated: false })
  }
  usage = (id: string, usage: AgentTokenUsage) => {
    this.patch(id, { usage: addTokenUsage(this.state(id).usage ?? undefined, usage) })
  }
  toolsAllowed = () => {
    const settings = this.deps.settings()
    return !!settings && localModelCapabilities(settings).tools !== false
  }
  prepare(run: WorkspaceStartedRun): WorkspaceRunPreparation {
    const id = run.binding.conversationId
    this.runs.set(id, run)
    this.patch(id, { error: null, truncated: false })
    return {
      project: run.binding.project,
      skills: [],
      dispose: () => {
        if (this.runs.get(id) === run) this.runs.delete(id)
      },
    }
  }
  transport(id: string, client: WorkspacePanelClient, purpose?: 'writer') {
    let settled = Promise.resolve()
    const subscriptions = new Set<{
      requestId: string | null
      listener(chunk: IpcStreamChunk): void
    }>()
    const finish = () => {
      for (const entry of [...subscriptions])
        if (entry.requestId) entry.listener({ requestId: entry.requestId, type: 'done' })
    }
    const transport = createIpcTransport<AiSettings>({
      getSettings: () => {
        const settings = this.deps.settings()
        if (!settings) throw new Error('AI settings unavailable')
        return settings
      },
      onStream: (listener) => {
        const entry = { requestId: null as string | null, listener }
        subscriptions.add(entry)
        const off = client.subscribe((event) => {
          if (event.kind === 'closed') {
            finish()
            return
          }
          const run = this.runs.get(id)
          if (
            event.kind === 'stream' &&
            run &&
            event.run.bindingId === run.binding.bindingId &&
            event.run.runId === run.binding.runId &&
            event.run.conversationId === id
          )
            listener(event.chunk)
        })
        return () => {
          subscriptions.delete(entry)
          off()
        }
      },
      start: ({ settings: _settings, ...request }) => {
        const run = this.runs.get(id)
        if (!run || run.signal.aborted) throw new Error('Workspace execution stopped')
        for (const entry of subscriptions) entry.requestId = request.requestId
        // Main reads the stored selection again and checks this exact live run.
        settled = client.stream(run.binding, { ...request, ...(purpose ? { purpose } : {}) })
        return settled
      },
      cancel: (requestId) => {
        const run = this.runs.get(id)
        if (run) void client.cancelStream(run.binding, requestId).catch(() => {})
      },
      unknownErrorText: this.deps.unknownError,
      errorText: (cause) => aiRunErrorText(this.deps.lang(), cause),
      contextLimit: localContextLimit,
      confirmContext: (warning) =>
        this.deps.confirm(aiContextConfirmText(this.deps.lang(), warning)),
      onStreamError: (error) => {
        this.patch(id, { error })
        this.deps.onError?.(error)
      },
      onSubscriptionUsage: (usage) =>
        this.patch(id, {
          subscriptionUsage: mergeSubscriptionUsage(this.state(id).subscriptionUsage, usage),
        }),
      onInputTruncated: () => this.patch(id, { truncated: true }),
    })
    return { ...transport, settled: () => settled, dispose: finish }
  }
}
