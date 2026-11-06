import { useEffect, useMemo, useState, type DragEvent } from 'react'
import { createI18n, type Lang } from '@genoffice/i18n'
import { DocAttachButton, DriveAttachButton, docMentionActions } from '@genoffice/ui'
import { aiStrings } from '../../../../../docs/src/renderer/i18n/strings-ai'
import type { AttachmentAddResult } from '../../../../../docs/src/shared/ipc'
import type { WorkspacePanelClient } from './panel-client'
import { workspacePanelText } from '../i18n/strings-workspace-panel'
import type { WorkspacePanelController } from './panel-controller'

export type WorkspaceAttachmentBridge = Pick<
  WorkspacePanelClient,
  'pickAttachments' | 'addAttachments' | 'pasteAttachment'
> &
  Partial<Pick<WorkspacePanelClient, 'driveStatus' | 'pickDriveAttachments'>>

const translate = createI18n(aiStrings)

/** Async picks keep the conversation that opened them, including after a conversation switch. */
export function useWorkspaceAttachments(
  controller: WorkspacePanelController,
  bridge: WorkspaceAttachmentBridge | undefined,
  id: string | null,
  lang: Lang,
) {
  const [pending, setPending] = useState(0)
  const [notice, setNotice] = useState<{ id: string; text: string } | null>(null)
  const t = (key: keyof typeof aiStrings.zh, params?: Record<string, string | number>) =>
    translate(lang, key, params)
  // The panel may unmount while a native picker is open.
  const [lifetime] = useState(() => ({ live: true }))
  useEffect(() => {
    lifetime.live = true
    return () => {
      lifetime.live = false
    }
  }, [lifetime])
  const merge = (conversationId: string, result: AttachmentAddResult | null, overflow = false) => {
    if (!lifetime.live || !result) return
    const composer = controller.host.conversations.state(conversationId).composer
    const files = new Map(composer.attachments.map((file) => [file.path, file]))
    for (const file of result.accepted) files.set(file.path, file)
    controller.host.conversations.edit(conversationId, {
      composer: { ...composer, attachments: [...files.values()].slice(0, 20) },
    })
    setNotice({
      id: conversationId,
      text: [
        ...result.rejected,
        ...(overflow || files.size > 20 ? [workspacePanelText(lang, 'attachmentLimit')] : []),
      ].join('\n'),
    })
  }
  const attach = async (paths?: string[], clipboard?: File[]) => {
    if (!bridge || !id || pending) return
    const conversationId = id
    setPending((value) => value + 1)
    try {
      const result: AttachmentAddResult = { accepted: [], rejected: [] }
      const collect = (added: AttachmentAddResult | null) => {
        if (added) {
          result.accepted.push(...added.accepted)
          result.rejected.push(...added.rejected)
        }
      }
      if (!paths && !clipboard) collect(await bridge.pickAttachments(conversationId))
      if (paths?.length) collect(await bridge.addAttachments(conversationId, paths.slice(0, 20)))
      for (const file of clipboard ?? []) {
        const ext = (
          {
            'image/png': 'png',
            'image/jpeg': 'jpg',
            'image/gif': 'gif',
            'image/webp': 'webp',
          } as Record<string, string>
        )[file.type]
        if (!ext || file.size > 5 * 1024 * 1024) {
          result.rejected.push(t('aiImageReadFail', { name: file.name }))
          continue
        }
        collect(await bridge.pasteAttachment(conversationId, await file.arrayBuffer(), ext))
      }
      merge(conversationId, result, (paths?.length ?? 0) > 20)
    } catch {
      if (lifetime.live)
        setNotice({ id: conversationId, text: workspacePanelText(lang, 'selectionFailed') })
    } finally {
      if (lifetime.live) setPending((value) => value - 1)
    }
  }
  const driveBridge = useMemo(
    () =>
      bridge?.driveStatus && bridge.pickDriveAttachments && id
        ? {
            driveConnected: async () =>
              (await bridge.driveStatus!().catch(() => ({ connected: false }))).connected,
            onDriveConnectedChanged: (listener: (connected: boolean) => void) =>
              window.aiOfficeDrive?.onStateChanged((state) =>
                listener(state.state === 'connected'),
              ) ?? (() => {}),
            drivePickAttachments: async () => {
              setPending((value) => value + 1)
              try {
                return await bridge.pickDriveAttachments!(id).catch(() => ({
                  ok: false,
                  reason: 'failed',
                }))
              } finally {
                if (lifetime.live) setPending((value) => value - 1)
              }
            },
          }
        : undefined,
    [bridge, id, lifetime],
  )
  const attachments = id ? controller.host.conversations.state(id).composer.attachments : []
  return {
    header:
      bridge && id ? (
        <div className="workspace-ai-attachments">
          {attachments.map((file, index) => (
            <div key={file.path ?? index} className="workspace-ai-row">
              <span className="workspace-ai-grow" dir="auto" title={file.path}>
                {file.name}
              </span>
              <button
                type="button"
                aria-label={`${t('aiRemoveAttachmentTitle')} · ${file.name}`}
                onClick={() => {
                  const composer = controller.host.conversations.state(id).composer
                  controller.host.conversations.edit(id, {
                    composer: {
                      ...composer,
                      attachments: composer.attachments.filter((item) => item.path !== file.path),
                    },
                  })
                }}
              >
                ×
              </button>
            </div>
          ))}
          {notice?.id === id && notice.text && <p role="alert">{notice.text}</p>}
        </div>
      ) : null,
    controls:
      bridge && id ? (
        <>
          <button
            type="button"
            disabled={pending > 0 || attachments.length >= 20}
            aria-label={t('aiAttachTitle')}
            title={t('aiAttachTitle')}
            onClick={() => {
              void attach()
            }}
          >
            +
          </button>
          {driveBridge && (
            <DriveAttachButton
              lang={lang}
              driveBridge={driveBridge}
              disabled={pending > 0 || attachments.length >= 20}
              onPicked={async ({ paths, notes }) => {
                if (paths.length) await attach(paths)
                if (notes.length && lifetime.live)
                  setNotice((current) => ({
                    id,
                    text: [
                      ...(paths.length && current?.id === id && current.text ? [current.text] : []),
                      ...notes,
                    ].join('\n'),
                  }))
              }}
            />
          )}
          <DocAttachButton
            lang={lang}
            onPicked={async ({ paths, notes }) => {
              if (paths.length) await attach(paths)
              if (notes.length) setNotice({ id, text: notes.join('\n') })
            }}
          />
        </>
      ) : null,
    onDragOver: (event: DragEvent) => {
      if (bridge && id && event.dataTransfer.types.includes('Files')) event.preventDefault()
    },
    onDrop: (event: DragEvent) => {
      if (!bridge || !id || !event.dataTransfer.types.includes('Files')) return
      event.preventDefault()
      event.stopPropagation()
      const paths = Array.from(event.dataTransfer.files)
        .map((file) => window.aiOffice.pathForFile(file))
        .filter(Boolean)
      if (paths.length) void attach(paths)
    },
    onPasteFiles: (files: File[]) => {
      const paths = files.map((file) => window.aiOffice.pathForFile(file)).filter(Boolean)
      const clipboard = files.filter((file) => !window.aiOffice.pathForFile(file))
      void attach(paths, clipboard)
    },
    docs:
      bridge && id
        ? {
            lang,
            ...docMentionActions({
              lang,
              add: (paths) => bridge.addAttachments(id, paths),
              merge: (result) => merge(id, result),
            }),
            attachments: attachments.flatMap((file) =>
              file.path ? [{ path: file.path, name: file.name }] : [],
            ),
          }
        : undefined,
  }
}
