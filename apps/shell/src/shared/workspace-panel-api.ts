import type { WorkspaceReadRequest } from '@genoffice/agent-core'
import type {
  WorkspaceCapture,
  WorkspaceExecutionController,
  WorkspaceReferenceCapture,
  WorkspaceStartedRun,
} from '../main/workspace-ai/execution-controller'
import type {
  WorkspaceHistory,
  WorkspaceHistoryMessage,
  WorkspaceConversationHistory,
} from '../main/workspace-ai/conversation-history'
import type {
  WorkspaceAiOwner,
  WorkspaceAiRejection,
  WorkspaceAiStoppedRun,
  WorkspaceConversationOrigin,
} from './workspace-ai-api'
import type { DocumentTabKind } from './tabs-api'
import type { WorkspaceLayout, WorkspacePanelLayout } from './workspace-layout'

export interface WorkspacePanelDocument {
  id: string
  title: string
  kind: DocumentTabKind
  /** Metadata only; main resolves the live view again on capture. */
  available: boolean
  active?: boolean
  pane?: 'left' | 'right' | 'top' | 'bottom'
}

export interface WorkspacePanelSource {
  origin: WorkspaceConversationOrigin
  title: string
  projectTitle?: string
}

export interface WorkspacePanelCatalog {
  documents: WorkspacePanelDocument[]
  sources: WorkspacePanelSource[]
}

export interface WorkspacePanelSession {
  id: string
  owner: WorkspaceAiOwner
}

/** AbortSignal cannot cross Electron IPC. The client reconstructs it from stopped events. */
export type WorkspacePanelStartedRun = Omit<WorkspaceStartedRun, 'signal'> & { aborted: boolean }

export type WorkspacePanelRequest =
  | { kind: 'catalog' }
  | { kind: 'layout'; panel: WorkspacePanelLayout | null }
  | { kind: 'capture'; tabId: string }
  | { kind: 'reference'; captureId: string; scope: WorkspaceReadRequest['scope'] }
  | { kind: 'create' }
  | { kind: 'load'; origin: WorkspaceConversationOrigin }
  | { kind: 'append'; origin: WorkspaceConversationOrigin; message: WorkspaceHistoryMessage }
  | { kind: 'enqueue'; request: Parameters<WorkspaceExecutionController['enqueue']>[0] }
  | { kind: 'startNext' }
  | { kind: 'execute'; args: Parameters<WorkspaceExecutionController['execute']> }
  | { kind: 'verifyResponse'; args: Parameters<WorkspaceExecutionController['verifyResponse']> }
  | { kind: 'complete'; args: Parameters<WorkspaceExecutionController['complete']> }
  | { kind: 'cancel'; args: Parameters<WorkspaceExecutionController['cancel']> }
  | { kind: 'state'; runId: string }
  | { kind: 'rollbackForRun'; args: Parameters<WorkspaceExecutionController['rollbackForRun']> }
  | { kind: 'rollback'; args: Parameters<WorkspaceExecutionController['rollback']> }
  | { kind: 'disconnect' }

export interface WorkspacePanelValues {
  catalog: WorkspacePanelCatalog
  layout: WorkspaceLayout
  capture: WorkspaceCapture
  reference: Pick<WorkspaceReferenceCapture, 'id' | 'binding'>
  create: WorkspacePanelSource
  load: WorkspaceHistory
  append: Awaited<ReturnType<WorkspaceConversationHistory['append']>>
  enqueue: ReturnType<WorkspaceExecutionController['enqueue']>
  startNext: WorkspacePanelStartedRun | null
  execute: Awaited<ReturnType<WorkspaceExecutionController['execute']>>
  verifyResponse: string | null
  complete: void
  cancel: void
  state: ReturnType<WorkspaceExecutionController['state']>
  rollbackForRun: ReturnType<WorkspaceExecutionController['rollbackForRun']>
  rollback: void
  disconnect: void
}

export type WorkspacePanelReply<T> =
  { ok: true; value: T } | { ok: false; reason: WorkspaceAiRejection | 'service-error' }

export type WorkspacePanelEvent = { sessionId: string } & (
  | { kind: 'catalog-changed' }
  | { kind: 'history-changed'; origin: WorkspaceConversationOrigin }
  | {
      kind: 'history-rebound'
      previous: WorkspaceConversationOrigin
      origin: Extract<WorkspaceConversationOrigin, { kind: 'legacy-document' }>
    }
  | { kind: 'layout'; layout: WorkspaceLayout }
  | { kind: 'stopped'; run: WorkspaceAiStoppedRun }
  | { kind: 'closed' }
)

export interface WorkspacePanelApi {
  connect(): Promise<WorkspacePanelReply<WorkspacePanelSession>>
  request<R extends WorkspacePanelRequest>(
    sessionId: string,
    request: R,
  ): Promise<WorkspacePanelReply<WorkspacePanelValues[R['kind']]>>
  onEvent(handler: (event: WorkspacePanelEvent) => void): () => void
}

export const WORKSPACE_PANEL_CHANNELS = {
  connect: 'workspace-ai:panel-connect',
  request: 'workspace-ai:panel-request',
  event: 'workspace-ai:panel-event',
} as const

declare global {
  interface Window {
    aiOfficeWorkspacePanel?: WorkspacePanelApi
  }
}
