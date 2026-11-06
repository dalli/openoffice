import type { DocsWorkspaceAiEntry } from '../../shared/workspace-ai-entry'

/** A rejected/unavailable host keeps the existing panel; an accepted entry is never resent. */
export async function routeDocsWorkspaceEntry(
  request: ((entry: DocsWorkspaceAiEntry) => Promise<boolean | 'rejected'>) | undefined,
  entry: DocsWorkspaceAiEntry,
  current: () => boolean,
  local: () => void,
  accepted: () => void,
): Promise<void> {
  let handled = false
  try {
    handled = (await request?.(entry)) === true
  } catch {
    // Standalone editors do not register the shell-only handler.
  }
  if (current()) (handled ? accepted : local)()
}
