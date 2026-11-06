import type { Editor } from '@tiptap/core'
import { WorkspaceAdapterError } from '@genoffice/agent-core'
import type { WorkspaceEditorControls } from '@genoffice/agent-core/workspace-editor-host'
import type {
  WorkspaceEditQueueCommand,
  WorkspaceEditQueueSnapshot,
} from '@genoffice/agent-core/workspace-editor-protocol'
import { resolveQueue, type DocsEditQueueItem } from './edit-queue'

/** The editor retains every anchor. The shell receives only redacted, bounded labels. */
export function createDocsEditQueueControls(deps: {
  editor(): Editor
  items(): DocsEditQueueItem[]
  update(id: string, instruction: string): void
  remove(id: string): void
  clear(): void
  focus(id: string): void
  send(): Promise<boolean>
}): WorkspaceEditorControls & { changed(): void } {
  const listeners = new Set<() => void>()
  let signature = ''
  let version = crypto.randomUUID() as string
  let sending = false
  const snapshot = (revision: string): WorkspaceEditQueueSnapshot => {
    const items = resolveQueue(deps.editor(), deps.items()).map(({ item, target }) => ({
      id: item.qid,
      instruction: item.instruction,
      excerpt: target?.excerpt.slice(0, 160) ?? '',
      available: target !== null,
    }))
    const next = JSON.stringify([revision, items])
    if (signature !== next) {
      signature = next
      version = crypto.randomUUID()
    }
    return { version, items }
  }
  return {
    subscribe(handler) {
      listeners.add(handler)
      return () => {
        listeners.delete(handler)
      }
    },
    changed() {
      for (const listener of listeners) listener()
    },
    async editQueue(command: WorkspaceEditQueueCommand, revision, signal) {
      if (signal.aborted) throw new WorkspaceAdapterError('cancelled')
      if (sending) throw new WorkspaceAdapterError('busy')
      const current = snapshot(revision)
      if (command.kind === 'list') return current
      if (command.version !== current.version) throw new WorkspaceAdapterError('stale-document')
      if ('id' in command && !current.items.some((i) => i.id === command.id))
        throw new WorkspaceAdapterError('invalid-call')
      switch (command.kind) {
        case 'update':
          deps.update(command.id, command.instruction.trim())
          break
        case 'remove':
          deps.remove(command.id)
          break
        case 'clear':
          deps.clear()
          break
        case 'focus':
          deps.focus(command.id)
          break
        case 'send': {
          if (!current.items.some((i) => i.available))
            throw new WorkspaceAdapterError('invalid-call')
          sending = true
          try {
            if (!(await deps.send())) throw new WorkspaceAdapterError('invalid-call')
            return null
          } finally {
            sending = false
          }
        }
      }
      return snapshot(revision)
    },
  }
}
