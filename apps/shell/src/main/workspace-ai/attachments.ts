import { stat } from 'node:fs/promises'
import { assertFileSaved } from '@genoffice/electron-utils/unsaved-files'
import type { ChatAttachment } from '@genoffice/project-store'
import type {
  AttachmentAddResult,
  AttachmentImageResult,
  AttachmentMeta,
  AttachmentReadResult,
} from '../../../../docs/src/shared/ipc'
import { ATTACHMENT_IMAGE_EXTS } from '../../../../docs/src/shared/ipc'
import { WorkspaceExecutionError, type WorkspaceRunHandle } from './execution-controller'

export interface WorkspaceAttachmentServices {
  paste?(bytes: ArrayBuffer, ext: string): AttachmentAddResult
  pick(): Promise<AttachmentAddResult | null>
  add(paths: string[]): AttachmentAddResult
  read(path: string, offset: number, maxChars: number): Promise<AttachmentReadResult>
  image(path: string): AttachmentImageResult
}
interface Entry {
  run: WorkspaceRunHandle
  files: AttachmentMeta[]
  stamps: string[]
}
const fail = (): never => {
  throw new WorkspaceExecutionError('invalid-request')
}
const stamp = async (path: string): Promise<string> => {
  const s = await stat(path)
  if (!s.isFile()) fail()
  return `${s.dev}:${s.ino}:${s.size}:${s.mtimeMs}:${s.ctimeMs}`
}
const sameRun = (a: WorkspaceRunHandle, b: WorkspaceRunHandle) =>
  a.bindingId === b.bindingId &&
  a.runId === b.runId &&
  a.conversationId === b.conversationId &&
  a.owner.windowId === b.owner.windowId &&
  a.owner.panelContentsId === b.owner.panelContentsId &&
  a.owner.panelGeneration === b.owner.panelGeneration

/** Paths are granted by explicit attachment actions, then frozen for one main-owned run. */
export class WorkspaceAttachments {
  private readonly granted = new Map<string, Map<string, AttachmentMeta>>()
  private readonly runs = new Map<number, Entry>()
  private closed = false
  constructor(private readonly services: WorkspaceAttachmentServices) {}

  async add(conversationId: string, paths?: string[]): Promise<AttachmentAddResult | null> {
    if (this.closed) fail()
    const result = paths ? this.services.add(paths) : await this.services.pick()
    if (this.closed) fail()
    return this.accept(conversationId, result)
  }
  paste(conversationId: string, bytes: ArrayBuffer, ext: string): AttachmentAddResult | null {
    if (
      this.closed ||
      !this.services.paste ||
      bytes.byteLength > 5 * 1024 * 1024 ||
      !bytes.byteLength ||
      !ATTACHMENT_IMAGE_EXTS.has(ext)
    )
      fail()
    return this.accept(conversationId, this.services.paste!(bytes, ext))
  }
  private accept(
    conversationId: string,
    result: AttachmentAddResult | null,
  ): AttachmentAddResult | null {
    if (!result) return null
    const granted = this.granted.get(conversationId) ?? new Map<string, AttachmentMeta>()
    for (const file of result.accepted) granted.set(file.path, structuredClone(file))
    // Composer metadata is cheap, but a long-lived window still needs a bound.
    while (granted.size > 256) granted.delete(granted.keys().next().value!)
    this.granted.set(conversationId, granted)
    return structuredClone(result)
  }

  async bind(run: WorkspaceRunHandle, requested: readonly ChatAttachment[]): Promise<void> {
    if (this.closed || this.runs.has(run.bindingId) || requested.length > 20) fail()
    const granted = this.granted.get(run.conversationId)
    const files = requested.map((file) => {
      const accepted = file.path ? granted?.get(file.path) : undefined
      if (
        !accepted ||
        accepted.name !== file.name ||
        accepted.ext !== file.ext ||
        accepted.sizeBytes !== file.sizeBytes
      )
        return fail()
      return structuredClone(accepted)
    })
    if (new Set(files.map((file) => file.path)).size !== files.length) fail()
    const entry: Entry = { run: structuredClone(run), files, stamps: [] }
    this.runs.set(run.bindingId, entry)
    try {
      for (const file of files) {
        await assertFileSaved(file.path)
        entry.stamps.push(await stamp(file.path))
        if (this.entry(run) !== entry) fail()
      }
    } catch (error) {
      this.release(run)
      throw error
    }
  }

  private entry(run: WorkspaceRunHandle): Entry {
    const entry = this.runs.get(run.bindingId)
    if (this.closed || !entry || !sameRun(entry.run, run)) return fail()
    return entry
  }
  private async checked(run: WorkspaceRunHandle, index: number): Promise<AttachmentMeta> {
    const entry = this.entry(run)
    const file = entry.files[index]
    if (!file) return fail()
    await assertFileSaved(file.path)
    const current = await stamp(file.path)
    if (this.entry(run) !== entry || current !== entry.stamps[index]) fail()
    return file
  }
  async list(run: WorkspaceRunHandle): Promise<AttachmentMeta[]> {
    const entry = this.entry(run)
    for (let i = 0; i < entry.files.length; i++) await this.checked(run, i)
    if (this.entry(run) !== entry) fail()
    return structuredClone(entry.files)
  }
  async read(
    run: WorkspaceRunHandle,
    index: number,
    offset: number,
  ): Promise<AttachmentReadResult> {
    const file = await this.checked(run, index)
    const result = await this.services.read(file.path, offset, 24_000)
    await this.checked(run, index)
    if (!result.ok) fail()
    return result
  }
  async image(run: WorkspaceRunHandle, index: number): Promise<AttachmentImageResult> {
    const file = await this.checked(run, index)
    if (!ATTACHMENT_IMAGE_EXTS.has(file.ext)) fail()
    const result = this.services.image(file.path)
    await this.checked(run, index)
    if (!result.ok) fail()
    return result
  }
  release(run: WorkspaceRunHandle): void {
    const entry = this.runs.get(run.bindingId)
    if (entry && sameRun(entry.run, run)) this.runs.delete(run.bindingId)
  }
  dispose(): void {
    this.closed = true
    this.granted.clear()
    this.runs.clear()
  }
}
