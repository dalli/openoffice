import { afterEach, describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { addQueueAnchor, removeQueueAnchors } from '../src/renderer/editor/ai-queue-anchors'
import { createDocsEditQueueControls } from '../src/renderer/ai/workspace-edit-queue'
import type { DocsEditQueueItem } from '../src/renderer/ai/edit-queue'

const cleanups: (() => void)[] = []
afterEach(() => cleanups.splice(0).forEach((off) => off()))
function fixture() {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: [
        {
          type: 'docParagraph',
          content: [
            { type: 'text', text: 'Visible ' },
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
  cleanups.push(() => editor.destroy())
  addQueueAnchor(editor, 'q1', 1, editor.state.doc.content.size - 1)
  let items: DocsEditQueueItem[] = [
    { qid: 'q1', instruction: 'Polish this', capturedText: 'HIDDEN_SECRET' },
    { qid: 'gone', instruction: 'Deleted passage', capturedText: 'HIDDEN_SECRET' },
  ]
  const send = vi.fn(async () => {
    items = []
    return true
  })
  const focus = vi.fn()
  const controls = createDocsEditQueueControls({
    editor: () => editor,
    items: () => items,
    update: (id, instruction) => {
      items = items.map((i) => (i.qid === id ? { ...i, instruction } : i))
    },
    remove: (id) => {
      removeQueueAnchors(editor, [id])
      items = items.filter((i) => i.qid !== id)
    },
    clear: () => {
      items = []
    },
    focus,
    send,
  })
  const signal = new AbortController().signal
  const list = async (revision = '1') =>
    (await controls.editQueue({ kind: 'list' }, revision, signal))!
  return { editor, controls, signal, list, send, focus, items: () => items }
}

describe('editor-owned workspace edit queue controls', () => {
  it('returns current redacted labels, never an orphan snapshot or block indexes', async () => {
    const f = fixture()
    const before = f.editor.getJSON()
    const snapshot = await f.list()
    expect(snapshot.items[0]).toMatchObject({ id: 'q1', available: true })
    expect(snapshot.items[0]!.excerpt).toContain('{{hidden}}')
    expect(snapshot.items[1]).toMatchObject({ id: 'gone', excerpt: '', available: false })
    expect(JSON.stringify(snapshot)).not.toContain('HIDDEN_SECRET')
    expect(JSON.stringify(snapshot)).not.toContain('startIndex')
    expect(f.editor.getJSON()).toEqual(before)
    expect(f.send).not.toHaveBeenCalled()
  })
  it('edits, focuses, removes, and clears only a current queue without changing document content', async () => {
    const f = fixture()
    const before = f.editor.getJSON()
    let snapshot = await f.list()
    snapshot = (await f.controls.editQueue(
      { kind: 'update', version: snapshot.version, id: 'q1', instruction: '  Shorten  ' },
      '1',
      f.signal,
    ))!
    expect(snapshot.items[0]!.instruction).toBe('Shorten')
    await f.controls.editQueue(
      { kind: 'focus', version: snapshot.version, id: 'q1' },
      '1',
      f.signal,
    )
    expect(f.focus).toHaveBeenCalledWith('q1')
    snapshot = (await f.controls.editQueue(
      { kind: 'remove', version: snapshot.version, id: 'gone' },
      '1',
      f.signal,
    ))!
    expect(snapshot.items).toHaveLength(1)
    await f.controls.editQueue({ kind: 'clear', version: snapshot.version }, '1', f.signal)
    expect(f.items()).toEqual([])
    expect(f.editor.getJSON()).toEqual(before)
    expect(f.send).not.toHaveBeenCalled()
  })
  it('rejects every old card operation after a queue or document revision changes', async () => {
    const f = fixture()
    const snapshot = await f.list()
    for (const command of [
      { kind: 'send', version: snapshot.version },
      { kind: 'clear', version: snapshot.version },
      { kind: 'remove', version: snapshot.version, id: 'q1' },
      { kind: 'focus', version: snapshot.version, id: 'q1' },
      { kind: 'update', version: snapshot.version, id: 'q1', instruction: 'Wrong' },
    ] as const)
      await expect(f.controls.editQueue(command, '2', f.signal)).rejects.toMatchObject({
        kind: 'stale-document',
      })
    expect(f.items()).toHaveLength(2)
    expect(f.send).not.toHaveBeenCalled()
    expect(f.focus).not.toHaveBeenCalled()
  })
  it('does not consume on rejected transfer or cancellation and serializes send attempts', async () => {
    const f = fixture()
    const { version } = await f.list()
    f.send.mockResolvedValueOnce(false)
    await expect(
      f.controls.editQueue({ kind: 'send', version }, '1', f.signal),
    ).rejects.toMatchObject({ kind: 'invalid-call' })
    expect(f.items()).toHaveLength(2)
    const abort = new AbortController()
    abort.abort()
    await expect(
      f.controls.editQueue({ kind: 'send', version }, '1', abort.signal),
    ).rejects.toMatchObject({ kind: 'cancelled' })
    let resolve!: (value: boolean) => void
    f.send.mockImplementationOnce(
      () =>
        new Promise((done) => {
          resolve = done
        }),
    )
    const sending = f.controls.editQueue({ kind: 'send', version }, '1', f.signal)
    await expect(
      f.controls.editQueue({ kind: 'send', version }, '1', f.signal),
    ).rejects.toMatchObject({ kind: 'busy' })
    resolve(true)
    expect(await sending).toBeNull()
    expect(f.send).toHaveBeenCalledTimes(2)
  })
})
