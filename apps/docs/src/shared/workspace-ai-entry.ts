/** Private, user-initiated navigation from the Word editor to its window's AI host. */
export const DOCS_WORKSPACE_AI_ENTRY = 'docs:workspace-ai-entry'
/** Rejected batches retain their queue; only an unavailable host permits local fallback. */
export type DocsWorkspaceAiEntryResult = boolean | 'rejected'

export type DocsWorkspaceAiEntry =
  | { kind: 'open' | 'toggle' | 'queue' }
  | { kind: 'preset'; text: string; autoRun: boolean }
  | {
      kind: 'edit-batch'
      text: string
      displayText: string
      documentId: string
      revision: string
    }

export function isDocsWorkspaceAiEntry(value: unknown): value is DocsWorkspaceAiEntry {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return false
  const entry = value as Record<string, unknown>
  if (entry.kind === 'edit-batch')
    return (
      Object.keys(entry).length === 5 &&
      ['text', 'displayText', 'documentId', 'revision'].every(
        (key) =>
          typeof entry[key] === 'string' &&
          entry[key].trim().length > 0 &&
          entry[key].length <= (key === 'documentId' || key === 'revision' ? 256 : 16_000),
      )
    )
  return entry.kind === 'open' || entry.kind === 'toggle' || entry.kind === 'queue'
    ? Object.keys(entry).length === 1
    : entry.kind === 'preset' &&
        Object.keys(entry).length === 3 &&
        typeof entry.text === 'string' &&
        entry.text.trim().length > 0 &&
        entry.text.length <= 64_000 &&
        typeof entry.autoRun === 'boolean'
}
