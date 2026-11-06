import { describe, expect, it } from 'vitest'
import { isWorkspacePanelRequest } from '../src/main/workspace-ai/panel-request'

describe('explicit workspace template requests', () => {
  const document = {
    windowId: 1,
    tabId: 'tab',
    documentId: 'document',
    generation: 1,
    kind: 'docs',
  }
  it.each(['info', 'preview', 'as-is'])(
    'admits the %s action through the panel request validator',
    (mode) => {
      expect(isWorkspacePanelRequest({ kind: 'template', document, mode })).toBe(true)
    },
  )
  it.each([
    { document: { ...document, windowId: 0 }, mode: 'preview' },
    { document: { ...document, generation: -1 }, mode: 'preview' },
    { document: { ...document, kind: 'pdf' }, mode: 'preview' },
    { document: { ...document, documentId: '' }, mode: 'preview' },
    { document, mode: 'save' },
  ])('rejects malformed or unsupported actions', (input) => {
    expect(isWorkspacePanelRequest({ kind: 'template', ...input })).toBe(false)
  })
})
