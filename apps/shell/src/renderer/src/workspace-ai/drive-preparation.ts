import {
  createDriveSkill,
  guardDriveContent,
  readAnswerOf,
  searchAnswerOf,
  type AgentMessage,
  type DriveSkill,
} from '@genoffice/agent-core'
import type { WorkspacePanelClient } from './panel-client'
import type { WorkspaceStartedRun } from '../../../main/workspace-ai/execution-controller'
import type { WorkspaceRunHost, WorkspaceRunPreparation } from './run-host'

type Bridge = Pick<WorkspacePanelClient, 'checkDrive' | 'searchDrive' | 'readDrive'>
interface Conversation {
  skill: DriveSkill
  enabled: boolean
  seen: boolean
  run: WorkspaceStartedRun | null
}

/** Keeps Drive guards with their conversation; only active run handles cross the preload. */
export class WorkspaceDrivePreparation {
  private readonly conversations = new Map<string, Conversation>()
  private closed = false
  constructor(
    private readonly bridge: Bridge,
    private readonly host: WorkspaceRunHost,
  ) {}
  private conversation(id: string): Conversation {
    const existing = this.conversations.get(id)
    if (existing) return existing
    const messages = (): readonly AgentMessage[] => [
      ...this.host.conversations.runtime(id).loop.messages,
      {
        role: 'tool',
        results: [
          ...this.host.conversations.state(id).messages.flatMap((m) =>
            (m.tools ?? [])
              .filter((t) => t.name === 'drive_read')
              .map((t) => ({
                id: `recorded-${m.seq}`,
                name: t.name,
                output: '',
                isError: t.isError,
              })),
          ),
          ...(entry.seen ? [{ id: 'drive-seen', name: 'drive_read', output: '' }] : []),
        ],
      },
    ]
    const entry: Conversation = { enabled: false, seen: false, run: null, skill: null! }
    const current = () => {
      const run = entry.run
      if (this.closed || !run || run.signal.aborted) throw new Error('Drive execution stopped')
      return run
    }
    const check = (run: WorkspaceStartedRun) => {
      if (current() !== run) throw new Error('Drive execution stopped')
    }
    entry.skill = createDriveSkill({
      enabled: () => entry.enabled && !!entry.run && !entry.run.signal.aborted && !this.closed,
      messages,
      search: async (query, limit) => {
        const run = current()
        const result = await this.bridge.searchDrive(run.binding, query, limit)
        check(run)
        return searchAnswerOf(result)
      },
      read: async (id, offset) => {
        const run = current()
        const result = readAnswerOf(await this.bridge.readDrive(run.binding, id, offset))
        check(run)
        if (result.ok) entry.seen = true
        return result
      },
    })
    this.conversations.set(id, entry)
    return entry
  }
  async prepare(run: WorkspaceStartedRun): Promise<WorkspaceRunPreparation> {
    if (this.closed) throw new Error('Drive preparation closed')
    const entry = this.conversation(run.binding.conversationId)
    if (entry.run) throw new Error('Drive already running')
    entry.run = run
    let disposed = false
    const check = () => {
      if (disposed || this.closed || run.signal.aborted || entry.run !== run)
        throw new Error('Drive execution stopped')
    }
    const beforeTurn = async () => {
      check()
      const status = await this.bridge.checkDrive(run.binding)
      check()
      entry.enabled = status.enabled
    }
    const dispose = () => {
      disposed = true
      if (entry.run === run) {
        entry.run = null
        entry.enabled = false
      }
    }
    try {
      await beforeTurn()
      return {
        project: run.binding.project,
        skills: [entry.skill],
        guard: (skill) => {
          const guarded = guardDriveContent(skill, entry.skill)
          return {
            ...guarded,
            get tools() {
              return guarded.tools
            },
            get systemPrompt() {
              return guarded.systemPrompt
            },
            executeTool: async (call, signal) => {
              check()
              if (signal?.aborted) throw new Error('Drive execution stopped')
              const result = await guarded.executeTool(call, signal)
              check()
              if (signal?.aborted) throw new Error('Drive execution stopped')
              return result
            },
          }
        },
        beforeTurn,
        dispose,
      }
    } catch (error) {
      dispose()
      throw error
    }
  }
  dispose() {
    this.closed = true
    for (const entry of this.conversations.values()) {
      entry.run = null
      entry.enabled = false
    }
    this.conversations.clear()
  }
}
