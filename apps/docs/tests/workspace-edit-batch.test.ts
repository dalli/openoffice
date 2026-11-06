import { afterEach, describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { addQueueAnchor } from '../src/renderer/editor/ai-queue-anchors'
import { createDocsWorkspaceAdapter } from '../src/renderer/ai/workspace-adapter'
import { handoffDocsEditBatch } from '../src/renderer/ai/workspace-edit-batch'
import type { DocsEditQueueItem } from '../src/renderer/ai/edit-queue'

const cleanups: Array<() => void> = []
afterEach(() =>
  cleanups
    .splice(0)
    .reverse()
    .forEach((off) => off()),
)
function fixture() {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        { type: 'docParagraph', content: [{ type: 'text', text: 'First passage' }] },
        {
          type: 'docParagraph',
          content: [
            { type: 'text', text: 'Last ' },
            {
              type: 'text',
              text: 'HIDDEN_SECRET',
              marks: [{ type: 'redaction', attrs: { label: 'hidden' } }],
            },
          ],
        },
      ],
    },
  })
  const adapter = createDocsWorkspaceAdapter({
    getEditor: () => editor,
    getDocumentId: () => 'fixture-document',
    getNumIds: () => ({ bullet: null, ordered: null }),
  })
  cleanups.push(
    () => editor.destroy(),
    () => adapter.dispose(),
  )
  let pos = 0
  editor.state.doc.forEach((node, _offset, index) => {
    addQueueAnchor(editor, `q${index}`, pos + 1, pos + node.nodeSize - 1)
    pos += node.nodeSize
  })
  const items: DocsEditQueueItem[] = [
    { qid: 'q0', instruction: 'Shorten first', capturedText: 'First passage' },
    { qid: 'q1', instruction: 'Polish last', capturedText: 'Last HIDDEN_SECRET' },
    { qid: 'orphan', instruction: 'Gone', capturedText: 'HIDDEN_SECRET' },
  ]
  const consume = vi.fn()
  return {
    editor,
    items,
    consume,
    deps: {
      capture: () => adapter.capture(),
      editor: () => editor,
      items: () => items,
      current: () => true,
      header: (count: number) => `${count} edits`,
      consume,
    },
  }
}

describe('Word edit batch transfer', () => {
  it('resolves migrated anchors, masks excerpts, omits orphan instructions, and consumes only after acknowledgement', async () => {
    const f = fixture()
    f.editor.commands.insertContentAt(1, 'Prefixed ')
    let acknowledge!: (accepted: boolean) => void
    const request = vi.fn(
      (_entry: import('../src/shared/workspace-ai-entry').DocsWorkspaceAiEntry) =>
        new Promise<boolean>((resolve) => {
          acknowledge = resolve
        }),
    )
    const pending = handoffDocsEditBatch({ ...f.deps, request })
    await vi.waitFor(() => expect(request).toHaveBeenCalledOnce())
    const payload = request.mock.calls[0]![0] as unknown as { text: string; displayText: string }
    expect(payload.text).toContain('First passage')
    expect(payload.text).not.toContain('Prefixed')
    expect(payload.text.indexOf('Polish last')).toBeLessThan(payload.text.indexOf('Shorten first'))
    expect(payload.text).not.toContain('Gone')
    expect(JSON.stringify(payload)).not.toContain('HIDDEN_SECRET')
    expect(payload.displayText).toContain('2 edits')
    expect(f.consume).not.toHaveBeenCalled()
    f.items[0]!.instruction = 'Revised while waiting'
    f.items.push({ qid: 'new', instruction: 'New request', capturedText: 'new' })
    acknowledge(true)
    expect(await pending).toBe(true)
    expect(f.consume).toHaveBeenCalledWith(['q1', 'orphan'])
  })

  it('retains the queue on rejection or an unavailable host', async () => {
    const f = fixture()
    expect(await handoffDocsEditBatch(f.deps)).toBe(false)
    for (const request of [
      async () => false,
      async () => {
        throw new Error("No handler registered for 'docs:workspace-ai-entry'")
      },
    ])
      expect(await handoffDocsEditBatch({ ...f.deps, request })).toBe(false)
    for (const request of [
      async () => 'rejected' as const,
      async () => {
        throw new Error('Closed')
      },
    ])
      expect(await handoffDocsEditBatch({ ...f.deps, request })).toBe(true)
    expect(f.consume).not.toHaveBeenCalled()
  })

  it('ignores late replies after document replacement and suppresses a local resend', async () => {
    const f = fixture()
    let current = true
    const request = vi.fn(async () => {
      current = false
      return false
    })
    expect(await handoffDocsEditBatch({ ...f.deps, request, current: () => current })).toBe(true)
    expect(f.consume).not.toHaveBeenCalled()
  })
})
