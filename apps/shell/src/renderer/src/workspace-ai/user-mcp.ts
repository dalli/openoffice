import {
  createMcpWindowState,
  createUserMcpSkill,
  guardOutbound,
  guardReceivedContent,
  isMcpToolName,
  type AgentMessage,
  type AgentSkill,
  type McpToolsBridgeLike,
  type McpWindowState,
  type OutboundBridgeLike,
} from '@genoffice/agent-core'
import { userMcpText } from '@genoffice/ui'
import type { Lang } from '@genoffice/i18n'
import type { WorkspaceStartedRun } from '../../../main/workspace-ai/execution-controller'
import type { WorkspaceRunHost, WorkspaceRunPreparation } from './run-host'

interface ConversationMcp {
  state: McpWindowState
  guarded: AgentSkill
  current: AgentSkill | null
  active: AbortController | null
  off(): void
}

/** Reuses the existing MCP client and guards, with an explicit owner for each conversation. */
export class WorkspaceUserMcp {
  private closed = false
  private readonly conversations = new Map<string, ConversationMcp>()

  constructor(
    private readonly bridge: McpToolsBridgeLike,
    private readonly outbound: OutboundBridgeLike,
    private readonly host: WorkspaceRunHost,
    private readonly lang: () => Lang,
  ) {}

  state(id: string): McpWindowState {
    return this.conversation(id).state
  }
  usableServers(id: string): ReadonlySet<string> {
    return this.state(id).usableServers()
  }
  private conversation(id: string): ConversationMcp {
    if (this.closed) throw new Error('Workspace MCP closed')
    const existing = this.conversations.get(id)
    if (existing) return existing
    const saved = this.host.conversations.state(id).mcp
    const base = createMcpWindowState(
      {
        list: () => this.bridge.list(),
        cancel: (id) => this.bridge.cancel(id),
        retry: (server) => this.bridge.retry(server),
        openSettings: (server) => this.bridge.openSettings(server),
        onChanged: (listener) => this.bridge.onChanged(listener),
        call: async (server, tool, input, callId) => {
          const active = entry.active
          await state.checkInternetTools()
          const definition = state
            .servers()
            .find((s) => s.name === server)
            ?.tools.find((t) => t.name === tool)
          if (
            this.closed ||
            !active ||
            active.signal.aborted ||
            entry.active !== active ||
            !state.usableServers().has(server) ||
            !definition ||
            definition.pending ||
            definition.unusable ||
            definition.permission === 'off'
          )
            return { ok: false, error: 'MCP call is no longer available.' }
          return this.bridge.call(server, tool, input, callId)
        },
      },
      this.outbound,
    )
    for (const name of saved.disabledServers) base.setServerOffInChat(name, true)
    for (const name of saved.alwaysAllowedTools) base.allowInChat(name)
    const messages = () => this.host.conversations.runtime(id).loop.messages
    const state: McpWindowState = {
      ...base,
      permissions: this.host.permissions(id),
      // This owner lives exactly as long as the conversation. A failed run may empty
      // the loop's history; that does not create a new conversation or reset its choices.
      followMessages() {},
      checkConversation() {},
      allowInChat: (name) => {
        base.allowInChat(name)
        const mcp = this.host.conversations.state(id).mcp
        if (!mcp.alwaysAllowedTools.includes(name))
          this.host.conversations.edit(id, {
            mcp: { ...mcp, alwaysAllowedTools: [...mcp.alwaysAllowedTools, name] },
          })
      },
    }
    state.setBlockedStepLabel((server, tool) =>
      userMcpText(this.lang(), 'blockedStep', { server, tool }),
    )
    const entry: ConversationMcp = {
      state,
      current: null,
      active: null,
      guarded: null!,
      off: () => {},
    }
    const delegate: AgentSkill = {
      id: 'workspace-guarded',
      get systemPrompt() {
        return entry.current?.systemPrompt ?? ''
      },
      get tools() {
        return entry.current?.tools ?? []
      },
      buildContext: () => entry.current?.buildContext?.() ?? '',
      executeTool: async (call, signal) => {
        if (this.closed || signal?.aborted || !entry.current)
          throw new Error('Workspace MCP execution stopped')
        const result = await entry.current.executeTool(call, signal)
        if (this.closed || signal?.aborted) throw new Error('Workspace MCP execution stopped')
        return result
      },
      verifyResponse: (text, calls) => entry.current?.verifyResponse?.(text, calls) ?? null,
    }
    const offOutbound: Array<() => void> = []
    const guardMessages = (): readonly AgentMessage[] => [
      ...messages(),
      // Restored model history intentionally contains only role/text. The guard alone
      // also reads recorded success metadata; no old tool body is sent to the model.
      {
        role: 'tool',
        results: this.host.conversations.state(id).messages.flatMap((message) =>
          (message.tools ?? [])
            .filter((tool) => isMcpToolName(tool.name))
            .map((tool, index) => ({
              id: `recorded-${message.seq}-${index}`,
              name: tool.name,
              output: '',
              isError: tool.isError,
            })),
        ),
      },
    ]
    entry.guarded = guardOutbound(
      guardReceivedContent(delegate, { keepsInside: state.keepsInside, messages: guardMessages }),
      {
        internetToolsEnabled: () => this.outbound.internetToolsEnabled(),
        confirmAddresses: (tool, addresses) => this.outbound.confirmAddresses(tool, addresses),
        onInternetToolsChanged: (listener) => {
          const off = this.outbound.onInternetToolsChanged?.(listener) ?? (() => {})
          offOutbound.push(off)
          return off
        },
      },
      messages,
    )
    const offState = state.subscribe(() => {
      if (this.closed) return
      const mcp = this.host.conversations.state(id).mcp
      const disabledServers = [...state.serversOffInChat()]
      if (JSON.stringify(mcp.disabledServers) !== JSON.stringify(disabledServers))
        this.host.conversations.edit(id, { mcp: { ...mcp, disabledServers } })
    })
    entry.off = () => {
      entry.active?.abort()
      state.permissions.denyAll()
      offState()
      state.dispose()
      for (const off of offOutbound) off()
    }
    this.conversations.set(id, entry)
    return entry
  }

  async prepare(run: WorkspaceStartedRun): Promise<WorkspaceRunPreparation> {
    const entry = this.conversation(run.binding.conversationId)
    if (entry.active) throw new Error('Workspace MCP already running')
    const active = new AbortController()
    const signal = AbortSignal.any([active.signal, run.signal])
    const stop = () => active.abort()
    run.signal.addEventListener('abort', stop, { once: true })
    if (run.signal.aborted) stop()
    entry.active = active
    const check = () => {
      if (this.closed || signal.aborted) throw new Error('Workspace MCP execution stopped')
    }
    const dispose = () => {
      active.abort()
      run.signal.removeEventListener('abort', stop)
      if (entry.active === active) {
        entry.active = null
        entry.current = null
      }
    }
    try {
      check()
      await entry.state.refresh()
      check()
      const capability = createUserMcpSkill(
        entry.state,
        () => this.host.conversations.runtime(run.binding.conversationId).loop.messages,
      )
      return {
        project: run.binding.project,
        skills: [capability],
        guard: (skill) => {
          check()
          entry.current = skill
          return {
            id: entry.guarded.id,
            get systemPrompt() {
              return entry.guarded.systemPrompt
            },
            get tools() {
              return entry.guarded.tools
            },
            buildContext: () => {
              check()
              return entry.guarded.buildContext?.() ?? ''
            },
            verifyResponse: (text, calls) => {
              check()
              return entry.guarded.verifyResponse?.(text, calls) ?? null
            },
            executeTool: async (call, callerSignal) => {
              check()
              const combined = callerSignal ? AbortSignal.any([signal, callerSignal]) : signal
              if (combined.aborted) throw new Error('Workspace MCP execution stopped')
              const result = await entry.guarded.executeTool(call, combined)
              check()
              return result
            },
          }
        },
        dispose,
      }
    } catch (error) {
      dispose()
      throw error
    }
  }

  dispose() {
    if (this.closed) return
    this.closed = true
    for (const entry of this.conversations.values()) entry.off()
    this.conversations.clear()
  }
}
