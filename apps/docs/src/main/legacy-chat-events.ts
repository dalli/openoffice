import type { WebContents } from 'electron'

export interface LegacyChatRef {
  projectId: string
  chatId: string
}
export type LegacyChatEvent =
  | { kind: 'resolved'; sender: WebContents; ref: LegacyChatRef; previous?: LegacyChatRef }
  | { kind: 'appended'; ref: LegacyChatRef }
  | { kind: 'renamed' }

const listeners = new Set<(event: LegacyChatEvent) => void>()

/** Main-only metadata notifications. They never include transcript or document contents. */
export function onLegacyChatEvent(listener: (event: LegacyChatEvent) => void): () => void {
  listeners.add(listener)
  return () => {
    listeners.delete(listener)
  }
}

export function emitLegacyChatEvent(event: LegacyChatEvent): void {
  for (const listener of listeners) {
    try {
      listener(event)
    } catch {
      /* An observer cannot change an existing chat operation. */
    }
  }
}
