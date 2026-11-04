import type { DocumentTabKind } from './tabs-api'

/** Main-owned run contracts. The shared panel API is not registered yet. */
export interface WorkspaceAiOwner {
  readonly windowId: number
  readonly panelContentsId: number
  readonly panelGeneration: number
}

/** A path or the currently focused tab is never an execution identity. */
export interface WorkspaceDocumentIdentity {
  readonly windowId: number
  readonly tabId: string
  readonly documentId: string
  readonly kind: DocumentTabKind
  readonly generation: number
}

export interface WorkspaceSelectionSnapshot {
  /** Opaque handle owned and revalidated by this editor generation. */
  readonly id: string
  readonly revision: string
}

export interface WorkspaceBoundDocument {
  readonly document: WorkspaceDocumentIdentity
  readonly revision: string
  readonly selection: WorkspaceSelectionSnapshot | null
}

export interface WorkspaceReference extends WorkspaceBoundDocument {
  /** Explicitly acquired current, redacted content; never a disk fallback. */
  readonly snapshotId: string
}

export interface WorkspaceProjectBinding {
  readonly id: string
  /** The prepared instruction/memory context version for this request. */
  readonly contextVersion: string
}

export interface WorkspaceAiRunRequest {
  readonly owner: WorkspaceAiOwner
  readonly conversationId: string
  readonly runId: string
  readonly target: WorkspaceBoundDocument
  readonly references: readonly WorkspaceReference[]
  readonly project: WorkspaceProjectBinding | null
}

/** Main-issued sequence prevents late results matching a reused, evicted run id. */
export interface WorkspaceAiBoundRun extends WorkspaceAiRunRequest {
  readonly bindingId: number
}

/** Document tools only; provider, filesystem, network and MCP tools keep their existing guards. */
export interface WorkspaceDocumentTool {
  readonly name: string
  readonly access: 'read' | 'write'
}

/** Registered by main from a known editor view, never supplied by the AI panel. */
export interface WorkspaceDocumentEndpoint {
  readonly document: WorkspaceDocumentIdentity
  readonly contentsId: number
  readonly revision: string
  readonly readOnly: boolean
  readonly tools: readonly WorkspaceDocumentTool[]
}

export interface WorkspaceAiToolRequest {
  readonly owner: WorkspaceAiOwner
  readonly conversationId: string
  readonly runId: string
  readonly bindingId: number
  readonly callId: string
  readonly document: WorkspaceDocumentIdentity
  readonly tool: string
}

/** A single-use dispatch ticket. It deliberately exposes no arbitrary IPC channel or JavaScript. */
export interface WorkspaceAiToolTicket extends WorkspaceAiToolRequest {
  readonly contentsId: number
  readonly access: 'read' | 'write'
  readonly revision: string
  readonly selection: WorkspaceSelectionSnapshot | null
  readonly snapshotId: string | null
}

export interface WorkspaceAiToolReply {
  readonly owner: WorkspaceAiOwner
  readonly conversationId: string
  readonly runId: string
  readonly bindingId: number
  readonly callId: string
  readonly document: WorkspaceDocumentIdentity
  readonly tool: string
  readonly revision: string
  readonly mutated: boolean
}

export type WorkspaceAiRejection =
  | 'wrong-owner'
  | 'invalid-request'
  | 'duplicate-run'
  | 'queue-full'
  | 'missing-document'
  | 'document-changed'
  | 'run-not-active'
  | 'run-finished'
  | 'call-in-flight'
  | 'duplicate-call'
  | 'unknown-tool'
  | 'read-only'
  | 'unbound-document'
  | 'wrong-response'
  | 'timeout'
  | 'cancelled'
  | 'document-closed'
  | 'editor-replaced'
  | 'panel-closed'
  | 'editor-error'
  | 'unsupported-read'

export type WorkspaceAiResult<T> =
  | { readonly ok: true; readonly value: T }
  | { readonly ok: false; readonly reason: WorkspaceAiRejection }

export type WorkspaceAiRunState =
  | { readonly status: 'queued' | 'running' | 'completed' }
  | { readonly status: 'stopped'; readonly reason: WorkspaceAiRejection }

export interface WorkspaceAiStoppedRun {
  readonly conversationId: string
  readonly runId: string
  readonly bindingId: number
  readonly reason: WorkspaceAiRejection
}

/** A conversation owns its origin; selecting a tab never rewrites this reference. */
export type WorkspaceConversationOrigin =
  | { readonly kind: 'legacy-document'; readonly projectId: string; readonly chatId: string }
  | { readonly kind: 'project'; readonly projectId: string; readonly chatId: string }
  | {
      readonly kind: 'window-memory'
      readonly windowSessionId: string
      readonly conversationId: string
    }
  | { readonly kind: 'workspace'; readonly conversationId: string }
