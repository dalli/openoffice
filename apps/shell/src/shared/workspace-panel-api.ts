import type { DocsWorkspaceAiEntry } from '../../../docs/src/shared/workspace-ai-entry'
import type {
  WorkspaceEditQueueCommand,
  WorkspaceEditQueueSnapshot,
} from '@genoffice/agent-core/workspace-editor-protocol'
import type { DriveAttachResult } from './drive-api'
import type { WorkspaceDriveStatus } from '../main/workspace-ai/drive'
import type { WorkspaceStreamRequest } from '../main/workspace-ai/streams'
import type { AiStreamChunk } from '@genoffice/ai-provider'
import type { ChatAttachment, ChatSkillChip } from '@genoffice/project-store'
import type {
  AttachmentAddResult,
  AttachmentMeta,
  AttachmentReadResult,
  AttachmentImageResult,
} from '../../../docs/src/shared/ipc'
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
  WorkspaceProjectBinding,
  WorkspaceDocumentIdentity,
} from './workspace-ai-api'
import type { WorkspaceTemplateMode, WorkspaceTemplateSnapshot } from '@genoffice/agent-core'
import type {
  WorkspacePreparedProject,
  WorkspaceProjectContext,
  WorkspaceProjectTool,
} from './workspace-project-api'
import type { WorkspaceRunHandle } from '../main/workspace-ai/execution-controller'
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
  | {
      kind: 'open-template'
      conversationId: string
      targetTabId?: string
      input: import('@genoffice/electron-utils/user-skills-channels').UserSkillTemplateOpenInput
    }
  | { kind: 'edit-queue'; document: WorkspaceDocumentIdentity; command: WorkspaceEditQueueCommand }
  | { kind: 'entry-ack'; entryId: string; accepted: boolean; conversationId?: string }
  | { kind: 'entries-ready'; ready: boolean }
  | { kind: 'template'; document: WorkspaceDocumentIdentity; mode: WorkspaceTemplateMode }
  | { kind: 'stream'; run: WorkspaceRunHandle; request: WorkspaceStreamRequest }
  | { kind: 'stream-cancel'; run: WorkspaceRunHandle; requestId: string }
  | { kind: 'drive-status' }
  | { kind: 'drive-pick'; conversationId: string }
  | { kind: 'drive-check'; run: WorkspaceRunHandle }
  | { kind: 'drive-search'; run: WorkspaceRunHandle; query: string; limit: number }
  | { kind: 'drive-read'; run: WorkspaceRunHandle; fileId: string; offset: number }
  | { kind: 'attachment-paste'; conversationId: string; bytes: ArrayBuffer; ext: string }
  | { kind: 'attachment-pick'; conversationId: string }
  | { kind: 'attachment-add'; conversationId: string; paths: string[] }
  | { kind: 'attachment-list'; run: WorkspaceRunHandle }
  | { kind: 'attachment-read'; run: WorkspaceRunHandle; index: number; offset: number }
  | { kind: 'attachment-image'; run: WorkspaceRunHandle; index: number }
  | { kind: 'catalog' }
  | { kind: 'layout'; panel: WorkspacePanelLayout | null }
  | { kind: 'capture'; tabId: string }
  | { kind: 'refresh'; captureId: string }
  | { kind: 'reference'; captureId: string; scope: WorkspaceReadRequest['scope'] }
  | { kind: 'create'; projectId?: string }
  | { kind: 'project-prepare'; conversationId: string; projectId: string }
  | { kind: 'project-discard'; conversationId: string; binding: WorkspaceProjectBinding }
  | { kind: 'project-context'; run: WorkspaceRunHandle }
  | {
      kind: 'project-tool'
      run: WorkspaceRunHandle
      method: WorkspaceProjectTool
      args: Record<string, unknown>
    }
  | { kind: 'load'; origin: WorkspaceConversationOrigin }
  | { kind: 'append'; origin: WorkspaceConversationOrigin; message: WorkspaceHistoryMessage }
  | {
      kind: 'enqueue'
      request: Parameters<WorkspaceExecutionController['enqueue']>[0]
      attachments?: ChatAttachment[]
    }
  | { kind: 'startNext' }
  | { kind: 'writer-update'; args: Parameters<WorkspaceExecutionController['updateWriter']> }
  | { kind: 'execute'; args: Parameters<WorkspaceExecutionController['execute']> }
  | { kind: 'verifyResponse'; args: Parameters<WorkspaceExecutionController['verifyResponse']> }
  | { kind: 'complete'; args: Parameters<WorkspaceExecutionController['complete']> }
  | { kind: 'cancel'; args: Parameters<WorkspaceExecutionController['cancel']> }
  | { kind: 'state'; runId: string }
  | { kind: 'rollbackForRun'; args: Parameters<WorkspaceExecutionController['rollbackForRun']> }
  | { kind: 'rollback'; args: Parameters<WorkspaceExecutionController['rollback']> }
  | { kind: 'disconnect' }

export interface WorkspacePanelValues {
  'open-template': unknown
  'edit-queue': WorkspaceEditQueueSnapshot | null
  'entry-ack': boolean
  'entries-ready': void
  template: WorkspaceTemplateSnapshot | null
  stream: void
  'stream-cancel': void
  'drive-status': WorkspaceDriveStatus
  'drive-pick': DriveAttachResult
  'drive-check': WorkspaceDriveStatus
  'drive-search': unknown
  'drive-read': unknown
  'attachment-paste': AttachmentAddResult | null
  'attachment-pick': AttachmentAddResult | null
  'attachment-add': AttachmentAddResult | null
  'attachment-list': AttachmentMeta[]
  'attachment-read': AttachmentReadResult
  'attachment-image': AttachmentImageResult
  catalog: WorkspacePanelCatalog
  layout: WorkspaceLayout
  capture: WorkspaceCapture
  refresh: WorkspaceCapture
  reference: Pick<WorkspaceReferenceCapture, 'id' | 'binding'>
  create: WorkspacePanelSource
  'project-prepare': WorkspacePreparedProject
  'project-discard': void
  'project-context': WorkspaceProjectContext | null
  'project-tool': unknown
  load: WorkspaceHistory
  append: Awaited<ReturnType<WorkspaceConversationHistory['append']>>
  enqueue: ReturnType<WorkspaceExecutionController['enqueue']>
  startNext: WorkspacePanelStartedRun | null
  'writer-update': void
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

/** Template requests originate in main's one-use handoff store, never in editor messages. */
export type WorkspacePanelEntry =
  | DocsWorkspaceAiEntry
  | {
      kind: 'template-handoff'
      id: string
      text: string
      chips: ChatSkillChip[]
      projectId?: string
    }

export type WorkspacePanelEvent = { sessionId: string } & (
  | { kind: 'edit-queue-changed'; tabId: string }
  | { kind: 'stream'; run: WorkspaceRunHandle; chunk: AiStreamChunk }
  | { kind: 'entry'; entry: WorkspacePanelEntry; capture: WorkspaceCapture; entryId?: string }
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
