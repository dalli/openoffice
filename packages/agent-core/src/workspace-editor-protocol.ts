import type {
  WorkspaceAdapterErrorKind,
  WorkspaceAdapterState,
  WorkspacePrepareRequest,
  WorkspaceReadRequest,
  WorkspaceReadSnapshot,
} from './workspace-adapter'
import type { ExecutedToolCall } from './skill'
import type { AgentToolCall, AgentToolDef, ToolExecution } from './types'

/** Private main-to-editor protocol. It is not an editor-to-editor or panel API. */
export const WORKSPACE_EDITOR_CHANNELS = {
  request: 'workspace-ai:editor-request',
  reply: 'workspace-ai:editor-reply',
  cancel: 'workspace-ai:editor-cancel',
} as const

export interface WorkspaceEditorSession {
  readonly id: string
  readonly runId: string
  readonly state: WorkspaceAdapterState
  readonly systemPrompt: string
  readonly context: string
  readonly tools: readonly AgentToolDef[]
  readonly toolAccess: Readonly<Record<string, 'read' | 'write'>>
}

export type WorkspaceEditorOperation =
  | { readonly kind: 'capture' }
  | {
      readonly kind: 'read'
      readonly state: WorkspaceAdapterState
      readonly read: WorkspaceReadRequest
    }
  | { readonly kind: 'prepare'; readonly request: WorkspacePrepareRequest }
  | {
      readonly kind: 'execute'
      readonly sessionId: string
      readonly runId: string
      readonly call: AgentToolCall
    }
  | {
      readonly kind: 'verify'
      readonly sessionId: string
      readonly runId: string
      readonly text: string
      readonly executed: readonly ExecutedToolCall[]
    }
  | { readonly kind: 'release'; readonly sessionId: string; readonly runId: string }

export interface WorkspaceEditorRequest {
  readonly id: string
  readonly operation: WorkspaceEditorOperation
}

export interface WorkspaceEditorValues {
  capture: WorkspaceAdapterState
  read: WorkspaceReadSnapshot
  prepare: WorkspaceEditorSession
  execute: { result: ToolExecution; state: WorkspaceAdapterState }
  verify: string | null
  release: null
}

export type WorkspaceEditorReply =
  | {
      readonly id: string
      readonly ok: true
      readonly value: WorkspaceEditorValues[keyof WorkspaceEditorValues]
    }
  | {
      readonly id: string
      readonly ok: false
      readonly error: WorkspaceAdapterErrorKind | 'failed'
    }

export interface WorkspaceEditorPort {
  onRequest(handler: (request: WorkspaceEditorRequest) => void): () => void
  onCancel(handler: (id: string) => void): () => void
  reply(reply: WorkspaceEditorReply): void
}

/** Structural Electron subset keeps this entry point safe in sandboxed preloads. */
interface EditorIpc {
  on(channel: string, listener: (event: unknown, value: unknown) => void): unknown
  removeListener(channel: string, listener: (event: unknown, value: unknown) => void): unknown
  send(channel: string, value: unknown): void
}

export function workspaceEditorPort(ipc: EditorIpc): WorkspaceEditorPort {
  const subscribe = <T>(channel: string, handler: (value: T) => void) => {
    const listener = (_event: unknown, value: unknown) => handler(value as T)
    ipc.on(channel, listener)
    return () => {
      ipc.removeListener(channel, listener)
    }
  }
  return {
    onRequest: (handler) => subscribe(WORKSPACE_EDITOR_CHANNELS.request, handler),
    onCancel: (handler) => subscribe(WORKSPACE_EDITOR_CHANNELS.cancel, handler),
    reply: (reply) => ipc.send(WORKSPACE_EDITOR_CHANNELS.reply, reply),
  }
}

declare global {
  interface Window {
    aiOfficeWorkspaceEditor?: WorkspaceEditorPort
  }
}
