import { WorkspaceAdapterError, type WorkspaceWriterUpdate } from '@genoffice/agent-core'
import type { AiDocWriter, DocWriteResult } from './doc-writer'

interface ActiveWriter {
  id: string
  ready: Promise<void>
  open(): void
  close(): void
  progress: ((text: string) => void) | null
  resolve: ((result: DocWriteResult) => void) | null
  settled: boolean
}

/** A single editor-owned draft. Only its matching live tool can receive output. */
export class WorkspaceDocWriter {
  private readonly waiting = new Set<() => void>()
  private active: ActiveWriter | null = null

  begin(id: string): AiDocWriter {
    if (this.active) throw new WorkspaceAdapterError('busy')
    let open!: () => void
    let close!: () => void
    const ready = new Promise<void>((resolve, reject) => {
      open = resolve
      close = () => reject(new WorkspaceAdapterError('released'))
    })
    void ready.catch(() => {})
    const active = {
      id,
      ready,
      open,
      close,
      progress: null,
      resolve: null,
      settled: false,
    } as ActiveWriter
    this.active = active
    for (const notify of [...this.waiting]) notify()
    return {
      write: (_spec, progress, signal) =>
        new Promise<DocWriteResult>((resolve) => {
          const done = (result: DocWriteResult) => {
            signal?.removeEventListener('abort', abort)
            active.resolve = null
            active.settled = true
            resolve(result)
          }
          const abort = () => done({ ok: false, error: 'cancelled' })
          active.progress = progress
          active.resolve = done
          if (signal?.aborted) abort()
          else signal?.addEventListener('abort', abort, { once: true })
          active.open()
        }),
    }
  }

  async update(id: string, update: WorkspaceWriterUpdate, signal?: AbortSignal): Promise<void> {
    if (!this.active && update.kind === 'ready')
      await new Promise<void>((resolve, reject) => {
        const cleanup = () => {
          this.waiting.delete(begin)
          signal?.removeEventListener('abort', abort)
        }
        const begin = () => {
          cleanup()
          resolve()
        }
        const abort = () => {
          cleanup()
          reject(new WorkspaceAdapterError('cancelled'))
        }
        this.waiting.add(begin)
        if (signal?.aborted) abort()
        else signal?.addEventListener('abort', abort, { once: true })
      })
    const active = this.active
    if (!active || active.id !== id || signal?.aborted)
      throw new WorkspaceAdapterError('invalid-call')
    await active.ready
    if (this.active !== active || active.settled || signal?.aborted)
      throw new WorkspaceAdapterError('released')
    if (update.kind === 'progress') active.progress?.(update.text)
    if (update.kind === 'result') active.resolve?.({ ...update, html: update.text })
  }

  finish(): void {
    const active = this.active
    if (!active) return
    active.resolve?.({ ok: false, error: 'released' })
    active.close()
    this.active = null
  }
}
