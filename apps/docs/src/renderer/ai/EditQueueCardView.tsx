import React, { useState } from 'react'
import type { WorkspaceEditQueueSnapshot } from '@genoffice/agent-core/workspace-editor-protocol'
import './edit-queue.css'

export interface EditQueueLabels {
  title: string
  hint: string
  orphan: string
  edit: string
  remove: string
  discard: string
  cancel: string
  send(count: number): string
  confirm(count: number): string
}
interface Props {
  items: WorkspaceEditQueueSnapshot['items']
  labels: EditQueueLabels
  busy: boolean
  onEditInstruction(id: string, instruction: string): void
  onRemove(id: string): void
  onDiscardAll(): void
  onSend(): void
  onFocus(id: string): void
}

/** beyond this the list starts collapsed so it never swallows the transcript */
const AUTO_COLLAPSE_FROM = 4

export function EditQueueCardView({
  items,
  labels,
  busy,
  onEditInstruction,
  onRemove,
  onDiscardAll,
  onSend,
  onFocus,
}: Props): React.JSX.Element | null {
  /** null = follow the length-based default; set once the user clicks the chevron */
  const [manualFold, setManualFold] = useState<boolean | null>(null)
  const [editingQid, setEditingQid] = useState<string | null>(null)
  const [draft, setDraft] = useState('')
  const [confirmDiscard, setConfirmDiscard] = useState(false)

  if (items.length === 0) return null
  const folded = manualFold ?? items.length >= AUTO_COLLAPSE_FROM
  const liveCount = items.filter((item) => item.available).length

  return (
    <div className="ai-queue" role="group" aria-label={labels.title}>
      <div className="ai-queue-head">
        <span className="ai-queue-title">{labels.title}</span>
        <span className="ai-queue-count">
          {items.length}/{10}
        </span>
        <button
          className={`ai-queue-fold${folded ? ' folded' : ''}`}
          onClick={() => setManualFold(!folded)}
          aria-expanded={!folded}
          aria-label={labels.title}
        >
          <svg width="14" height="14" viewBox="0 0 16 16" aria-hidden>
            <path
              d="M4 6l4 4 4-4"
              fill="none"
              stroke="currentColor"
              strokeWidth="1.5"
              strokeLinecap="round"
              strokeLinejoin="round"
            />
          </svg>
        </button>
      </div>
      {!folded && (
        <>
          <div className="ai-queue-hint">{labels.hint}</div>
          <ol className="ai-queue-list">
            {items.map((item, i) => {
              const excerpt = item.excerpt
              return (
                <li
                  key={item.id}
                  className={`ai-queue-row${item.available ? '' : ' stale'}`}
                  data-tip={item.available ? undefined : labels.orphan}
                >
                  <span className="ai-queue-ord">{i + 1}</span>
                  {editingQid === item.id ? (
                    <input
                      className="ai-queue-edit-input"
                      value={draft}
                      disabled={busy}
                      aria-label={labels.edit}
                      autoFocus
                      maxLength={500}
                      onChange={(e) => setDraft(e.target.value)}
                      onBlur={() => {
                        const next = draft.trim()
                        if (!busy && next) onEditInstruction(item.id, next)
                        setEditingQid(null)
                      }}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter' && !e.nativeEvent.isComposing) {
                          e.preventDefault()
                          e.currentTarget.blur()
                        } else if (e.key === 'Escape') {
                          e.preventDefault()
                          setEditingQid(null)
                        }
                      }}
                    />
                  ) : (
                    <button
                      type="button"
                      className="ai-queue-text"
                      title={item.instruction}
                      disabled={busy || !item.available}
                      onClick={() => onFocus(item.id)}
                    >
                      {excerpt && (
                        <span className="ai-queue-target">
                          {excerpt.length > 24 ? `${excerpt.slice(0, 24)}…` : excerpt}
                        </span>
                      )}
                      {item.instruction}
                    </button>
                  )}
                  <span className="ai-queue-row-actions" onClick={(e) => e.stopPropagation()}>
                    <button
                      className="ai-queue-row-btn"
                      data-tip={labels.edit}
                      aria-label={labels.edit}
                      disabled={busy}
                      onClick={() => {
                        setDraft(item.instruction)
                        setEditingQid(item.id)
                      }}
                    >
                      <svg width="14" height="14" viewBox="0 0 16 16" aria-hidden>
                        <path
                          d="M11.2 2.8l2 2L5.6 12.4l-2.6.6.6-2.6z"
                          fill="none"
                          stroke="currentColor"
                          strokeWidth="1.3"
                          strokeLinejoin="round"
                        />
                      </svg>
                    </button>
                    <button
                      className="ai-queue-row-btn"
                      data-tip={labels.remove}
                      aria-label={labels.remove}
                      disabled={busy}
                      onClick={() => onRemove(item.id)}
                    >
                      <svg width="14" height="14" viewBox="0 0 16 16" aria-hidden>
                        <path
                          d="M4 4l8 8M12 4l-8 8"
                          fill="none"
                          stroke="currentColor"
                          strokeWidth="1.3"
                          strokeLinecap="round"
                        />
                      </svg>
                    </button>
                  </span>
                </li>
              )
            })}
          </ol>
        </>
      )}
      <div className="ai-queue-foot">
        {confirmDiscard ? (
          <>
            <span className="ai-queue-confirm">{labels.confirm(items.length)}</span>
            <button className="ai-queue-discard" onClick={() => setConfirmDiscard(false)}>
              {labels.cancel}
            </button>
            <button
              className="ai-queue-send"
              disabled={busy}
              onClick={() => {
                setConfirmDiscard(false)
                onDiscardAll()
              }}
            >
              {labels.discard}
            </button>
          </>
        ) : (
          <>
            <button
              className="ai-queue-discard"
              disabled={busy}
              onClick={() => setConfirmDiscard(true)}
            >
              {labels.discard}
            </button>
            <button className="ai-queue-send" disabled={busy || liveCount === 0} onClick={onSend}>
              {labels.send(liveCount)}
            </button>
          </>
        )}
      </div>
    </div>
  )
}
