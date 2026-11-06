import { describe, expect, it, vi } from 'vitest'
import { isDocsWorkspaceAiEntry } from '../src/shared/workspace-ai-entry'
import { routeDocsWorkspaceEntry } from '../src/renderer/ai/workspace-entry'

describe('Word workspace entry', () => {
  it('accepts only bounded navigation and explicit preset requests', () => {
    expect(isDocsWorkspaceAiEntry({ kind: 'toggle' })).toBe(true)
    expect(isDocsWorkspaceAiEntry({ kind: 'preset', text: 'Polish', autoRun: true })).toBe(true)
    const batch = {
      kind: 'edit-batch',
      text: 'instructions',
      displayText: '2 edits',
      documentId: 'doc',
      revision: '1',
    }
    expect(isDocsWorkspaceAiEntry(batch)).toBe(true)
    for (const extra of [
      { documentId: '' },
      { revision: 1 },
      { text: 'x'.repeat(16_001) },
      { displayText: ' ' },
      { tabId: 'other' },
    ])
      expect(isDocsWorkspaceAiEntry({ ...batch, ...extra })).toBe(false)
    for (const value of [
      null,
      [],
      { kind: 'open', tabId: 'other' },
      { kind: 'preset', text: ' ', autoRun: true },
      { kind: 'preset', text: 'x'.repeat(64_001), autoRun: true },
      { kind: 'preset', text: 'Polish', autoRun: 'true' },
    ])
      expect(isDocsWorkspaceAiEntry(value)).toBe(false)
  })

  it('falls back only before acceptance and ignores a reply after document replacement', async () => {
    const local = vi.fn()
    const accepted = vi.fn()
    const entry = { kind: 'open' as const }
    for (const request of [
      undefined,
      vi.fn(async () => false),
      vi.fn(async () => {
        throw new Error('No shell')
      }),
    ])
      await routeDocsWorkspaceEntry(request, entry, () => true, local, accepted)
    expect(local).toHaveBeenCalledTimes(3)
    await routeDocsWorkspaceEntry(
      async () => true,
      entry,
      () => true,
      local,
      accepted,
    )
    expect(accepted).toHaveBeenCalledOnce()
    await routeDocsWorkspaceEntry(
      async () => false,
      entry,
      () => false,
      local,
      accepted,
    )
    expect(local).toHaveBeenCalledTimes(3)
  })
})
