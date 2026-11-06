import { createI18n, type Lang } from '@genoffice/i18n'
import { createFilesSkill } from '../../../../../docs/src/renderer/ai/files-skill-core'
import { aiStrings } from '../../../../../docs/src/renderer/i18n/strings-ai'
import { ATTACHMENT_IMAGE_EXTS } from '../../../../../docs/src/shared/ipc'
import type { WorkspaceStartedRun } from '../../../main/workspace-ai/execution-controller'
import type { WorkspacePanelClient } from './panel-client'
import type { WorkspaceComposer } from './conversations'
import type { WorkspaceRunPreparation } from './run-host'

const translate = createI18n(aiStrings)

/** Reuse the editor's attachment protocol; only this run's granted indices can be read. */
export async function prepareWorkspaceAttachments(
  client: Pick<WorkspacePanelClient, 'attachmentList' | 'attachmentRead' | 'attachmentImage'>,
  input: WorkspaceComposer,
  run: WorkspaceStartedRun,
  lang: () => Lang,
): Promise<WorkspaceRunPreparation> {
  let closed = false
  const result: WorkspaceRunPreparation = {
    project: run.binding.project,
    skills: [],
    images: [],
    dispose() {
      closed = true
      result.images = []
    },
  }
  if (!input.attachments.length) return result
  const check = () => {
    if (closed || run.signal.aborted) throw new Error('Attachment execution stopped')
  }
  check()
  const files = await client.attachmentList(run.binding)
  check()
  if (
    files.length !== input.attachments.length ||
    files.some((file, index) => {
      const original = input.attachments[index]!
      return (
        file.path !== original.path ||
        file.name !== original.name ||
        file.ext !== original.ext ||
        file.sizeBytes !== original.sizeBytes
      )
    })
  )
    throw new Error('Attachment selection changed')
  for (let index = 0; index < files.length; index++) {
    if (!ATTACHMENT_IMAGE_EXTS.has(files[index]!.ext)) continue
    const image = await client.attachmentImage(run.binding, index)
    check()
    if (!image.ok || !image.base64 || !image.mime) throw new Error('Attachment image unavailable')
    result.images!.push({ base64: image.base64, mime: image.mime })
  }
  result.skills.push(
    createFilesSkill(() => files, {
      text: (key, params) => translate(lang(), key as keyof typeof aiStrings.zh, params),
      read: async (path, offset) => {
        check()
        const index = files.findIndex((file) => file.path === path)
        if (index < 0) throw new Error('Attachment not granted')
        const text = await client.attachmentRead(
          run.binding,
          index,
          Math.max(0, Math.floor(offset)),
        )
        check()
        return text
      },
    }),
  )
  return result
}
