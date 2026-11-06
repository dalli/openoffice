import type { AttachmentMeta } from '../../shared/ipc'
import { t } from '../i18n/locale'
import { createFilesSkill as createCore } from './files-skill-core'

/** Existing editor bridge; the workspace supplies its own guarded run bridge to the core. */
export function createFilesSkill(getAttachments: () => AttachmentMeta[]) {
  return createCore(getAttachments, {
    read: (...args) => window.desktop.readAttachment(...args),
    text: t,
  })
}
