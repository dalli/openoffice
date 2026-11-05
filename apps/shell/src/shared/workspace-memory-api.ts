import type { ChatMessage } from '@genoffice/project-store'
import type { WorkspaceHistoryMessage } from '../main/workspace-ai/conversation-history'

export type WorkspaceMemoryRequest = { hostId: string; id: string } & (
  { kind: 'load' } | { kind: 'append'; message: WorkspaceHistoryMessage } | { kind: 'release' }
)
export interface WorkspaceMemoryHost {
  /** Claims the original transcript. Must refuse while a legacy send is pending or running. */
  load(): readonly ChatMessage[]
  append(message: WorkspaceHistoryMessage): void
  release(): void
}
export interface WorkspaceMemoryApi {
  register(host: WorkspaceMemoryHost): Promise<() => void>
}
export const WORKSPACE_MEMORY_CHANNELS = {
  register: 'workspace-ai:memory-register',
  unregister: 'workspace-ai:memory-unregister',
  request: 'workspace-ai:memory-request',
  reply: 'workspace-ai:memory-reply',
} as const
declare global {
  interface Window {
    aiOfficeWorkspaceMemory?: WorkspaceMemoryApi
  }
}
