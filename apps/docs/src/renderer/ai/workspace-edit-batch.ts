import type { Editor } from '@tiptap/core'
import type { WorkspaceAdapterState } from '@genoffice/agent-core'
import type { DocsWorkspaceAiEntry } from '../../shared/workspace-ai-entry'
import {
  buildQueueInstruction,
  buildQueueSummary,
  liveItems,
  resolveQueue,
  type DocsEditQueueItem,
} from './edit-queue'

/** Transfer ownership only after the shell has installed the original document's batch draft. */
export async function handoffDocsEditBatch(deps: {
  request?: (entry: DocsWorkspaceAiEntry) => Promise<boolean | 'rejected'>
  capture(): Promise<WorkspaceAdapterState>
  editor(): Editor
  items(): DocsEditQueueItem[]
  current(): boolean
  header(count: number): string
  consume(qids: string[]): void
  rejected?(): void
}): Promise<boolean> {
  if (!deps.request) return false
  try {
    const state = await deps.capture()
    if (!deps.current()) return true
    const items = deps.items().map((item) => ({ ...item }))
    const entries = liveItems(resolveQueue(deps.editor(), items))
    if (!entries.length) return false
    const accepted = await deps.request({
      kind: 'edit-batch',
      documentId: state.instanceId,
      revision: state.revision,
      text: buildQueueInstruction(entries),
      displayText: buildQueueSummary(deps.header(entries.length), entries),
    })
    if (!deps.current()) return true
    if (accepted === 'rejected') {
      deps.rejected?.()
      return true
    }
    if (!accepted) return false
    // A late acknowledgement must preserve annotations added or revised during the handoff.
    const unchanged = items.filter((item) =>
      deps
        .items()
        .some(
          (now) =>
            now.qid === item.qid &&
            now.instruction === item.instruction &&
            now.capturedText === item.capturedText,
        ),
    )
    deps.consume(unchanged.map((item) => item.qid))
    return true
  } catch (error) {
    if (
      error instanceof Error &&
      error.message.includes("No handler registered for 'docs:workspace-ai-entry'")
    )
      return false
    // An interrupted IPC reply does not prove the host failed to take ownership.
    if (deps.current()) deps.rejected?.()
    return true
  }
}
