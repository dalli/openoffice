import type { ChatMessage } from '@genoffice/project-store'
import type { DriveAttachResult } from '../../shared/drive-api'
import {
  WorkspaceExecutionError,
  type WorkspaceRunHandle,
  type WorkspaceStartedRun,
} from './execution-controller'

export interface WorkspaceDriveStatus {
  connected: boolean
  enabled: boolean
  allowed: boolean
}
export interface WorkspaceDriveServices {
  status(): WorkspaceDriveStatus
  search(query: string, limit: number): Promise<unknown>
  read(id: string, offset: number): Promise<unknown>
  pick(): Promise<DriveAttachResult>
  isDrivePath(path: string): boolean
}
interface Entry {
  run: WorkspaceStartedRun
  paths: readonly string[]
}
const fail = (): never => {
  throw new WorkspaceExecutionError('invalid-request')
}
const sameRun = (a: WorkspaceRunHandle, b: WorkspaceRunHandle) =>
  a.bindingId === b.bindingId &&
  a.runId === b.runId &&
  a.conversationId === b.conversationId &&
  a.owner.windowId === b.owner.windowId &&
  a.owner.panelContentsId === b.owner.panelContentsId &&
  a.owner.panelGeneration === b.owner.panelGeneration

/** The existing Drive service, restricted to live runs of this panel. No tokens or raw errors. */
export class WorkspaceDrive {
  private readonly runs = new Map<number, Entry>()
  private readonly seen = new Set<string>()
  private closed = false
  constructor(private readonly services: WorkspaceDriveServices) {}
  restore(id: string, messages: readonly ChatMessage[]) {
    if (
      messages.some(
        (m) =>
          m.tools?.some((t) => t.name === 'drive_read' && !t.isError) ||
          m.attachments?.some((file) => !!file.path && this.services.isDrivePath(file.path)),
      )
    )
      this.seen.add(id)
  }
  bind(run: WorkspaceStartedRun, paths: readonly string[]) {
    if (this.closed || run.signal.aborted || this.runs.has(run.binding.bindingId)) fail()
    this.runs.set(run.binding.bindingId, { run, paths: [...paths] })
    if (paths.some((path) => this.services.isDrivePath(path)))
      this.seen.add(run.binding.conversationId)
  }
  private entry(handle: WorkspaceRunHandle): Entry {
    const entry = this.runs.get(handle.bindingId)
    if (this.closed || !entry || entry.run.signal.aborted || !sameRun(entry.run.binding, handle))
      return fail()
    return entry
  }
  check(handle: WorkspaceRunHandle): WorkspaceDriveStatus {
    const entry = this.entry(handle)
    const status = this.services.status()
    if (
      !status.allowed &&
      (this.seen.has(handle.conversationId) ||
        entry.paths.some((p) => this.services.isDrivePath(p)))
    )
      fail()
    return status
  }
  async tool(
    handle: WorkspaceRunHandle,
    method: 'search' | 'read',
    text: string,
    offset: number,
  ): Promise<unknown> {
    const entry = this.entry(handle)
    if (!this.check(handle).enabled)
      return { ok: false, error: 'Google Drive is not available now.' }
    const result =
      method === 'search'
        ? await this.services.search(text, offset)
        : await this.services.read(text, offset)
    if (this.entry(handle) !== entry || !this.check(handle).enabled) fail()
    if (method === 'read' && (result as { ok?: unknown } | null)?.ok === true)
      this.seen.add(handle.conversationId)
    return result
  }
  release(handle: WorkspaceRunHandle) {
    const entry = this.runs.get(handle.bindingId)
    if (entry && sameRun(entry.run.binding, handle)) this.runs.delete(handle.bindingId)
  }
  dispose() {
    this.closed = true
    this.runs.clear()
    this.seen.clear()
  }
}
