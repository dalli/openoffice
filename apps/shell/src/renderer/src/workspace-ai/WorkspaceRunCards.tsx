import { useSyncExternalStore } from 'react'
import type { Lang } from '@genoffice/i18n'
import { McpPermissionPrompt } from '@genoffice/ui'
import type { WorkspaceRunHost } from './run-host'

/** Labels come from the panel's 21-language resources, just like its existing undo cards. */
export interface WorkspaceRunCardLabels {
  undo: string
  restoring: string
  reverted: string
  refused: string
}

/** Switching the visible conversation never unmounts its permission owner or denies a pending call. */
export function WorkspaceRunCards({
  host,
  conversationId,
  lang,
  labels,
}: {
  host: WorkspaceRunHost
  conversationId: string
  lang: Lang
  labels: WorkspaceRunCardLabels
}) {
  useSyncExternalStore(host.subscribe, host.version, host.version)
  const request = host.permissions(conversationId).requests()[0]
  return (
    <>
      {request && (
        <McpPermissionPrompt
          key={`${conversationId}:${request.id}`}
          lang={lang}
          request={request}
          onAnswer={(answer) => host.answerPermission(conversationId, request.id, answer)}
        />
      )}
      {host.rollbackCards(conversationId).map((card) => (
        <section key={card.id} className="ai-mcp-card" role="group" aria-label={card.label}>
          <p dir="auto">{card.label}</p>
          <button
            type="button"
            className="ai-mcp-card-btn"
            disabled={card.state !== 'applied' || host.busy()}
            onClick={() => {
              void host.rollback(conversationId, card.id)
            }}
          >
            {card.state === 'applied' ? labels.undo : labels[card.state]}
          </button>
        </section>
      ))}
    </>
  )
}
