import type { AppendProjectChatResult, ChatMessage, ProjectStore } from '@genoffice/project-store'
import type { WorkspaceConversationOrigin } from '../../shared/workspace-ai-api'
import type { ChatSender, ProjectChatStore } from '../projects/chat-store'

export interface WorkspaceHistory {
  readonly messages: readonly ChatMessage[]
  /** Retained for the existing foreign-record notice; never grants project trust. */
  readonly foreign: boolean
}

export type WorkspaceHistoryMessage = Omit<ChatMessage, 'seq' | 'ts'>

export interface WorkspaceMemoryHistorySource {
  load(): readonly ChatMessage[]
  append(message: WorkspaceHistoryMessage): void
}

/** Main-only router. Origins are references, never instructions to migrate or merge records. */
export class WorkspaceConversationHistory {
  private readonly memory = new Map<string, ChatMessage[]>()
  private readonly memorySources = new Map<string, WorkspaceMemoryHistorySource>()
  private readonly newProjectChats = new Map<string, ChatMessage[]>()

  constructor(
    private readonly local: ProjectStore,
    private readonly projects: Pick<ProjectChatStore, 'openChat' | 'append' | 'newChat'>,
    private readonly windowSessionId: string,
  ) {}

  createLocal(title = ''): WorkspaceConversationOrigin {
    const info = this.local.createWorkspaceChat(title)
    return { kind: 'workspace', conversationId: info.id }
  }

  async createProject(projectId: string): Promise<WorkspaceConversationOrigin> {
    const { chatId } = await this.projects.newChat(projectId)
    this.newProjectChats.set(JSON.stringify([projectId, chatId]), [])
    return { kind: 'project', projectId, chatId }
  }

  /** The home host keeps owning its ephemeral transcript; opening it never copies it to disk. */
  attachMemory(conversationId: string, source: WorkspaceMemoryHistorySource): () => void {
    if (this.memorySources.has(conversationId) || this.memory.has(conversationId)) {
      throw new Error('Memory conversation is already attached')
    }
    this.memorySources.set(conversationId, source)
    return () => {
      if (this.memorySources.get(conversationId) === source)
        this.memorySources.delete(conversationId)
    }
  }

  private assertMemory(
    origin: Extract<WorkspaceConversationOrigin, { kind: 'window-memory' }>,
  ): void {
    if (origin.windowSessionId !== this.windowSessionId) {
      throw new Error('Conversation belongs to another window session')
    }
  }

  async load(origin: WorkspaceConversationOrigin): Promise<WorkspaceHistory> {
    switch (origin.kind) {
      case 'legacy-document':
        return {
          messages: this.local.loadChat(origin.projectId, origin.chatId),
          foreign: false,
        }
      case 'workspace':
        return {
          messages: this.local.loadWorkspaceChat(origin.conversationId),
          foreign: false,
        }
      case 'project': {
        const result = await this.projects.openChat(origin)
        const key = JSON.stringify([origin.projectId, origin.chatId])
        if (result.status !== 'ok') {
          const opening = this.newProjectChats.get(key)
          if (opening) return { messages: structuredClone(opening), foreign: false }
          throw new Error('Project conversation is not available')
        }
        this.newProjectChats.delete(key)
        return { messages: result.messages, foreign: result.foreign }
      }
      case 'window-memory':
        this.assertMemory(origin)
        return {
          messages: structuredClone(
            this.memorySources.get(origin.conversationId)?.load() ??
              this.memory.get(origin.conversationId) ??
              [],
          ),
          foreign: false,
        }
    }
  }

  /** Receives an already sanitized message from the host, and a main-identified sender. */
  async append(
    origin: WorkspaceConversationOrigin,
    message: WorkspaceHistoryMessage,
    sender: ChatSender,
  ): Promise<'recorded' | 'memory' | AppendProjectChatResult['status']> {
    switch (origin.kind) {
      case 'legacy-document':
        this.local.appendChatMessage(origin.projectId, origin.chatId, message)
        return 'recorded'
      case 'workspace':
        this.local.appendWorkspaceChatMessage(origin.conversationId, message)
        return 'recorded'
      case 'project': {
        const result = await this.projects.append(
          { ...message, projectId: origin.projectId, chatId: origin.chatId },
          sender,
        )
        const key = JSON.stringify([origin.projectId, origin.chatId])
        if (result.status === 'pending') this.newProjectChats.get(key)?.push(result.message)
        else this.newProjectChats.delete(key)
        return result.status
      }
      case 'window-memory': {
        this.assertMemory(origin)
        const source = this.memorySources.get(origin.conversationId)
        if (source) {
          source.append(structuredClone(message))
          return 'memory'
        }
        const messages = this.memory.get(origin.conversationId) ?? []
        messages.push({
          ...structuredClone(message),
          seq: (messages.at(-1)?.seq ?? -1) + 1,
          ts: new Date().toISOString(),
        })
        this.memory.set(origin.conversationId, messages.slice(-200))
        return 'memory'
      }
    }
  }
}
