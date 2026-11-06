import { useEffect, useMemo, useRef, useState } from 'react'
import type { Lang } from '@genoffice/i18n'
import type {
  WorkspaceEditQueueCommand,
  WorkspaceEditQueueSnapshot,
} from '@genoffice/agent-core/workspace-editor-protocol'
import type { WorkspaceDocumentIdentity } from '../../../shared/workspace-ai-api'
import type { WorkspacePanelEvent } from '../../../shared/workspace-panel-api'
import { EditQueueCardView } from '../../../../../docs/src/renderer/ai/EditQueueCardView'
import { workspaceDocText } from './docs-actions'
import { useI18n } from '../locale'

export interface WorkspaceEditQueueBridge {
  editQueue(
    document: WorkspaceDocumentIdentity,
    command: WorkspaceEditQueueCommand,
  ): Promise<WorkspaceEditQueueSnapshot | null>
  subscribe(handler: (event: WorkspacePanelEvent) => void): () => void
}

/** Mounted per conversation/target. No queue copy is added to conversation or model history. */
export function WorkspaceEditQueue({
  bridge,
  document: input,
  busy,
  lang,
}: {
  bridge: WorkspaceEditQueueBridge
  document: WorkspaceDocumentIdentity
  busy: boolean
  lang: Lang
}) {
  const { windowId, tabId, kind, documentId, generation } = input
  const document = useMemo(
    () => ({ windowId, tabId, kind, documentId, generation }),
    [windowId, tabId, kind, documentId, generation],
  )
  const { t } = useI18n()
  const [snapshot, setSnapshot] = useState<WorkspaceEditQueueSnapshot | null>(null)
  const [pending, setPending] = useState(false)
  const [failed, setFailed] = useState(false)
  const [refresh, setRefresh] = useState(0)
  const lifecycle = useRef({ alive: false, serial: 0, acting: false }).current
  useEffect(() => {
    lifecycle.alive = true
    const off = bridge.subscribe((event) => {
      if (event.kind === 'edit-queue-changed' && event.tabId === document.tabId)
        setRefresh((n) => n + 1)
    })
    return () => {
      lifecycle.alive = false
      lifecycle.serial++
      off()
    }
  }, [bridge, document, lifecycle])
  useEffect(() => {
    if (busy || lifecycle.acting) return
    const epoch = ++lifecycle.serial
    setPending(true)
    void bridge
      .editQueue(document, { kind: 'list' })
      .then((value) => {
        if (lifecycle.alive && epoch === lifecycle.serial) {
          setSnapshot(value)
          setFailed(false)
        }
      })
      .catch(() => {
        if (lifecycle.alive && epoch === lifecycle.serial) setFailed(true)
      })
      .finally(() => {
        if (lifecycle.alive && epoch === lifecycle.serial) setPending(false)
      })
  }, [bridge, document, busy, refresh, lifecycle])

  const act = async (command: WorkspaceEditQueueCommand) => {
    if (busy || lifecycle.acting || pending || !snapshot) return
    const epoch = ++lifecycle.serial
    lifecycle.acting = true
    setPending(true)
    setFailed(false)
    try {
      const value = await bridge.editQueue(document, command)
      if (lifecycle.alive && epoch === lifecycle.serial) setSnapshot(value)
    } catch {
      if (lifecycle.alive && epoch === lifecycle.serial) setFailed(true)
    } finally {
      lifecycle.acting = false
      if (lifecycle.alive && epoch === lifecycle.serial) {
        setPending(false)
        setRefresh((n) => n + 1)
      }
    }
  }
  if (!snapshot?.items.length) return null
  const version = snapshot.version
  return (
    <div className="workspace-ai-edit-queue">
      {failed && <p role="alert">{workspaceDocText(lang, 'aiQueueTransferFailed')}</p>}
      <EditQueueCardView
        key={snapshot.version}
        items={snapshot.items}
        busy={busy || pending || failed}
        labels={{
          title: workspaceDocText(lang, 'aiQueueTitle'),
          hint: workspaceDocText(lang, 'aiQueueHint'),
          edit: workspaceDocText(lang, 'aiQueueRowEdit'),
          orphan: workspaceDocText(lang, 'aiQueueOrphan'),
          discard: workspaceDocText(lang, 'aiQueueDiscard'),
          remove: t('delete'),
          cancel: t('cancel'),
          send: (count) => workspaceDocText(lang, 'aiQueueSend', { count }),
          confirm: (count) => workspaceDocText(lang, 'aiQueueDiscardConfirm', { count }),
        }}
        onEditInstruction={(id, instruction) => {
          void act({ kind: 'update', version, id, instruction })
        }}
        onRemove={(id) => {
          void act({ kind: 'remove', version, id })
        }}
        onDiscardAll={() => {
          void act({ kind: 'clear', version })
        }}
        onFocus={(id) => {
          void act({ kind: 'focus', version, id })
        }}
        onSend={() => {
          void act({ kind: 'send', version })
        }}
      />
    </div>
  )
}
