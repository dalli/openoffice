import React from 'react'
import type { Editor } from '@tiptap/core'
import { useI18n } from '../i18n/locale'
import { resolveQueue, type DocsEditQueueItem } from './edit-queue'
import { EditQueueCardView } from './EditQueueCardView'

interface Props {
  items: DocsEditQueueItem[]
  editor: Editor
  busy: boolean
  onEditInstruction(id: string, instruction: string): void
  onRemove(id: string): void
  onDiscardAll(): void
  onSend(): void
  onFocus(id: string): void
}
export function EditQueueCard({ items, editor, ...props }: Props): React.JSX.Element | null {
  const { t } = useI18n()
  return (
    <EditQueueCardView
      {...props}
      items={resolveQueue(editor, items).map(({ item, target }) => ({
        id: item.qid,
        instruction: item.instruction,
        excerpt: target?.excerpt ?? item.capturedText,
        available: target !== null,
      }))}
      labels={{
        title: t('aiQueueTitle'),
        hint: t('aiQueueHint'),
        orphan: t('aiQueueOrphan'),
        edit: t('aiQueueRowEdit'),
        remove: t('ribbonGroupDelete'),
        discard: t('aiQueueDiscard'),
        cancel: t('appCancel'),
        send: (count) => t('aiQueueSend', { count }),
        confirm: (count) => t('aiQueueDiscardConfirm', { count }),
      }}
    />
  )
}
