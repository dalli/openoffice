import { afterEach, describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { closeHistory } from '@tiptap/pm/history'
import type { AgentToolCall, WorkspaceDocumentAdapter } from '@genoffice/agent-core'
import { buildExtensions } from '../src/renderer/editor/extensions'
import { createMarkdownSkill } from '../src/renderer/ai/markdown-skill'
import { createMarkdownWorkspaceAdapter } from '../src/renderer/ai/workspace-adapter'

const editors: Editor[] = []
const adapters: WorkspaceDocumentAdapter[] = []
afterEach(() => {
  for (const adapter of adapters.splice(0)) adapter.dispose()
  for (const editor of editors.splice(0)) editor.destroy()
})

function editorOf(markdown = '# Note\n\nalpha old text\n\nbeta untouched'): Editor {
  const editor = new Editor({
    extensions: buildExtensions({
      slashController: { onOpen() {}, onUpdate() {}, onKeyDown: () => false, onClose() {} },
      slashItems: () => [],
    }),
    content: '',
  })
  editor.commands.setContent(markdown, { contentType: 'markdown' })
  editors.push(editor)
  return editor
}

function positionOf(editor: Editor, needle: string): number {
  let found = -1
  editor.state.doc.descendants((node, position) => {
    if (found >= 0 || !node.isText) return
    const offset = node.text!.indexOf(needle)
    if (offset >= 0) found = position + offset
  })
  expect(found, needle).toBeGreaterThanOrEqual(0)
  return found
}

function select(editor: Editor, needle: string): void {
  const from = positionOf(editor, needle)
  editor.view.dispatch(
    editor.state.tr.setSelection(
      TextSelection.create(editor.state.doc, from, from + needle.length),
    ),
  )
}

function fixture(editor = editorOf(), initialFrontmatter = '') {
  let current: Editor | null = editor
  let documentId = 'document-A'
  let source = false
  let yaml = initialFrontmatter
  const metadataChanges = new Set<() => void>()
  const frontmatter = {
    read: () => yaml,
    write: (inner: string) => {
      yaml = inner
      for (const notify of metadataChanges) notify()
    },
  }
  const adapter = createMarkdownWorkspaceAdapter({
    getEditor: () => current,
    getDocumentId: () => documentId,
    isSourceMode: () => source,
    frontmatter,
    rollback: {
      capture: () => ({ doc: editor.state.doc, yaml }),
      restore: (value) => {
        const snapshot = value as { doc: typeof editor.state.doc; yaml: string }
        editor.commands.setContent(snapshot.doc.toJSON())
        frontmatter.write(snapshot.yaml)
      },
    },
    subscribeChanges: (handler) => {
      metadataChanges.add(handler)
      return () => {
        metadataChanges.delete(handler)
      }
    },
    beforeFirstWrite: () => editor.view.dispatch(closeHistory(editor.state.tr)),
  })
  adapters.push(adapter)
  return {
    editor,
    adapter,
    frontmatter,
    replaceEditor: (next: Editor | null) => {
      current = next
    },
    replaceDocument: () => {
      documentId = 'document-B'
    },
    sourceMode: () => {
      source = true
    },
  }
}

let callId = 0
function call(name: string, input: Record<string, unknown> = {}): AgentToolCall {
  return { id: `call-${++callId}`, name, input }
}
const ops = (...list: Record<string, unknown>[]) => call('apply_ops', { ops: list })
async function prepare(adapter: WorkspaceDocumentAdapter, mode: 'edit' | 'reference' = 'edit') {
  const state = await adapter.capture()
  return adapter.prepare({ runId: `run-${++callId}`, mode, state })
}

describe('Markdown workspace snapshots', () => {
  it('refreshes the pinned range through a rewrite and rollback without following the visible caret', async () => {
    const f = fixture()
    select(f.editor, 'old text')
    const pinned = await f.adapter.capture()
    const first = await f.adapter.prepare({ runId: 'first', mode: 'edit', state: pinned })
    select(f.editor, 'beta')
    await first.execute(
      ops({ op: 'replaceText', target: 'selection', find: 'old text', replace: 'new words' }),
    )
    first.release()
    const refreshed = await f.adapter.refresh(pinned)
    expect((await f.adapter.read(refreshed, { scope: 'selection' })).text).toContain('new words')
    await f.adapter.rollback!(first.rollbackId!, first.state)
    const restored = await f.adapter.refresh(refreshed)
    expect((await f.adapter.read(restored, { scope: 'selection' })).text).toContain('old text')
    const second = await f.adapter.prepare({ runId: 'second', mode: 'edit', state: restored })
    await second.execute(
      ops({ op: 'replaceText', target: 'selection', find: 'old text', replace: 'continued' }),
    )
    second.release()
    expect(f.editor.state.doc.textContent).toContain('alpha continued')
    expect(f.editor.state.doc.textContent).toContain('beta untouched')
    f.editor.view.dispatch(f.editor.state.tr.insertText('manual ', 1))
    await expect(f.adapter.refresh(restored)).rejects.toMatchObject({ kind: 'stale-selection' })
  })

  it('restores the body and frontmatter together through the editor-owned token', async () => {
    const f = fixture(editorOf(), 'title: Original')
    const before = f.editor.getJSON()
    const session = await prepare(f.adapter)
    await session.execute(
      ops(
        { op: 'replaceText', target: { start: 1 }, find: 'old', replace: 'new' },
        { op: 'setFrontmatter', yaml: 'title: Changed' },
      ),
    )
    const last = session.state
    const id = session.rollbackId!
    session.release()
    await f.adapter.rollback!(id, last)
    expect(f.editor.getJSON()).toEqual(before)
    expect(f.frontmatter.read()).toBe('title: Original')
  })
  it('invalidates metadata snapshots as well as body edits without reading their text', async () => {
    const host = fixture()
    await host.adapter.capture()
    const changed = vi.fn()
    const off = host.adapter.subscribeChanges(changed)
    host.frontmatter.write('title: Changed')
    host.editor.view.dispatch(host.editor.state.tr.insertText('manual ', 1))
    expect(changed).toHaveBeenCalledTimes(2)
    expect(changed.mock.calls).toEqual([[], []])
    off()
    host.frontmatter.write('title: Later')
    expect(changed).toHaveBeenCalledTimes(2)
  })
  it('reads current unsaved structure and metadata with explicit truncation', async () => {
    const { adapter, editor } = fixture(
      editorOf('# Draft\n\nunsaved current material'),
      'title: Draft',
    )
    const state = await adapter.capture()
    const snapshot = await adapter.read(state, { scope: 'document' })
    expect(snapshot.text).toContain('# Draft')
    expect(snapshot.text).toContain('unsaved current material')
    expect(snapshot.text).toContain('title: Draft')
    expect(snapshot.truncated).toBe(false)
    const small = await adapter.read(state, { scope: 'document', maxChars: 12 })
    expect(small.text).toHaveLength(12)
    expect(small.truncated).toBe(true)
    expect(editor.getMarkdown()).toContain('unsaved current material')
  })

  it('reads a captured selection after the caret moves, including private placeholders', async () => {
    const { adapter, editor } = fixture(editorOf('call secret-phone now\n\nother public text'))
    select(editor, 'secret-phone')
    editor.commands.setRedaction('phone')
    select(editor, 'secret-phone')
    const state = await adapter.capture()
    editor.commands.setTextSelection(positionOf(editor, 'other'))
    const visible = editor.state.selection
    const selected = await adapter.read(state, { scope: 'selection' })
    const whole = await adapter.read(state, { scope: 'document' })
    const session = await adapter.prepare({ runId: 'private-read', mode: 'reference', state })
    const blocks = await session.execute(call('read_blocks', { startIndex: 0, endIndex: 1 }))
    for (const text of [selected.text, whole.text, session.context, blocks.result.output]) {
      expect(text).not.toContain('secret-phone')
      expect(text).toContain('{{phone}}')
    }
    expect(editor.state.selection.eq(visible)).toBe(true)
    expect(editor.getMarkdown()).toContain('secret-phone')
  })

  it('rejects a missing selection, source mode and unavailable editors without a disk fallback', async () => {
    const f = fixture()
    const state = await f.adapter.capture()
    await expect(f.adapter.read(state, { scope: 'selection' })).rejects.toMatchObject({
      kind: 'stale-selection',
    })
    f.sourceMode()
    await expect(f.adapter.capture()).rejects.toMatchObject({ kind: 'not-ready' })
    const missing = fixture()
    missing.replaceEditor(null)
    await expect(missing.adapter.capture()).rejects.toMatchObject({ kind: 'not-ready' })
  })

  it('rejects marked metadata before a snapshot or prepared context can expose it', async () => {
    const { adapter, frontmatter } = fixture(
      editorOf('public body'),
      'private: <span data-redaction="" data-label="key">secret-yaml</span>',
    )
    const state = await adapter.capture()
    await expect(adapter.read(state, { scope: 'document' })).rejects.toMatchObject({
      kind: 'unsafe-content',
    })
    await expect(
      adapter.prepare({ runId: 'unsafe-yaml', mode: 'edit', state }),
    ).rejects.toMatchObject({ kind: 'unsafe-content' })
    expect(frontmatter.read()).toContain('secret-yaml')
  })

  it('preserves the existing private-placeholder edit guard', async () => {
    const f = fixture(editorOf('call secret-phone now'))
    select(f.editor, 'secret-phone')
    f.editor.commands.setRedaction('phone')
    const before = f.editor.getJSON()
    const session = await prepare(f.adapter)
    const edited = await session.execute(
      ops({ op: 'replaceBlocks', target: { start: 0 }, markdown: 'call someone now' }),
    )
    expect(edited.result.isError).toBe(true)
    expect(edited.result.output).toContain('private placeholder')
    expect(edited.result.output).not.toContain('secret-phone')
    expect(f.editor.getJSON()).toEqual(before)
  })
})

describe('Markdown workspace document tools', () => {
  it('retains only local tools and refuses media, writing, saving and network calls', async () => {
    const { adapter, editor } = fixture()
    const session = await prepare(adapter)
    expect(session.tools.map((tool) => tool.name)).toEqual([
      'get_document_context',
      'read_blocks',
      'apply_ops',
      'read_frontmatter',
    ])
    const before = editor.getJSON()
    for (const name of [
      'write_document',
      'image_search',
      'insert_image',
      'generate_image',
      'generate_mindmap',
      'save',
      'read_url',
    ]) {
      await expect(session.execute(call(name))).rejects.toMatchObject({ kind: 'unknown-tool' })
    }
    const imageOp = await session.execute(
      ops({ op: 'insertImage', after: 0, src: 'https://example.invalid/private.png' }),
    )
    expect(imageOp.result.isError).toBe(true)
    const imageMarkdown = await session.execute(
      ops({
        op: 'insertContent',
        after: 0,
        markdown: '![photo](https://example.invalid/photo.png)',
      }),
    )
    expect(imageMarkdown.result.isError).toBe(true)
    expect(editor.getJSON()).toEqual(before)
  })

  it('matches the original skill for the same batch and preserves the same undo boundary', async () => {
    const legacy = editorOf()
    const f = fixture(editorOf(), 'title: Before')
    let legacyYaml = 'title: Before'
    const legacySkill = createMarkdownSkill(
      () => legacy,
      {
        read: () => legacyYaml,
        write: (value) => {
          legacyYaml = value
        },
      },
      () => false,
    )
    legacySkill.buildContext?.()
    legacy.view.dispatch(closeHistory(legacy.state.tr))
    const session = await prepare(f.adapter)
    const input = ops(
      { op: 'replaceText', target: { start: 1 }, find: 'old', replace: 'revised' },
      { op: 'setBlockType', target: { start: 2 }, type: 'heading', level: 2 },
      { op: 'setFrontmatter', yaml: 'title: Revised\ntags: [draft]' },
    )
    const original = await legacySkill.executeTool(input)
    const adapted = await session.execute(input)
    expect(adapted.result).toEqual(original)
    expect(f.editor.getJSON()).toEqual(legacy.getJSON())
    expect(f.editor.getMarkdown()).toBe(legacy.getMarkdown())
    expect(f.frontmatter.read()).toBe(legacyYaml)
    expect(f.editor.commands.undo()).toBe(legacy.commands.undo())
    expect(f.editor.getJSON()).toEqual(legacy.getJSON())
  })

  it('keeps inline formatting on the captured range while restoring the visible caret and typing marks', async () => {
    const f = fixture()
    select(f.editor, 'old')
    const state = await f.adapter.capture()
    f.editor.commands.setTextSelection(positionOf(f.editor, 'beta') + 2)
    f.editor.commands.setMark('italic')
    const visible = f.editor.state.selection
    const typingMarks = f.editor.state.storedMarks
    const session = await f.adapter.prepare({ runId: 'frozen-inline', mode: 'edit', state })
    expect(f.editor.state.selection.eq(visible)).toBe(true)
    expect(f.editor.state.storedMarks).toEqual(typingMarks)
    expect(session.context).toContain('## User selection (block 1)\nold')
    const styled = await session.execute(
      ops({ op: 'setStyle', target: 'selection', style: 'bold' }),
    )
    expect(styled.result.isError).toBeUndefined()
    expect(f.editor.getMarkdown()).toContain('alpha **old** text')
    expect(f.editor.getMarkdown()).not.toContain('**beta')
    expect(f.editor.state.selection.eq(visible)).toBe(true)
    expect(f.editor.state.storedMarks).toEqual(typingMarks)
    await session.execute(
      ops({ op: 'replaceText', target: 'selection', find: 'old', replace: 'longer replacement' }),
    )
    expect(f.editor.state.selection.from).toBe(positionOf(f.editor, 'beta') + 2)
    const context = await session.execute(call('get_document_context'))
    expect(context.result.output).toContain('## User selection (block 1)\nlonger replacement')
  })

  it('retains nested-list selection semantics instead of broadening to the whole top-level list', async () => {
    const legacy = editorOf('- first item\n- second item\n\noutside')
    const f = fixture(editorOf('- first item\n- second item\n\noutside'))
    select(legacy, 'first')
    select(f.editor, 'first')
    const original = createMarkdownSkill(
      () => legacy,
      undefined,
      () => false,
    )
    original.buildContext?.()
    const state = await f.adapter.capture()
    f.editor.commands.setTextSelection(positionOf(f.editor, 'outside'))
    const session = await f.adapter.prepare({ runId: 'nested', mode: 'edit', state })
    const input = ops({ op: 'setBlockType', target: 'selection', type: 'heading', level: 2 })
    const a = await original.executeTool(input)
    const b = await session.execute(input)
    expect(b.result).toEqual(a)
    expect(f.editor.getJSON()).toEqual(legacy.getJSON())
    expect(f.editor.getMarkdown()).toContain('## first item')
    expect(f.editor.getMarkdown()).toContain('- second item')
  })

  it('refuses a selection operation after the original selected range was removed', async () => {
    const f = fixture()
    select(f.editor, 'old')
    const session = await prepare(f.adapter)
    await session.execute(ops({ op: 'deleteBlocks', target: 'selection' }))
    const afterDelete = f.editor.getJSON()
    await expect(
      session.execute(ops({ op: 'setStyle', target: 'selection', style: 'bold' })),
    ).rejects.toMatchObject({ kind: 'stale-selection' })
    expect(f.editor.getJSON()).toEqual(afterDelete)
    expect(f.editor.getMarkdown()).toContain('beta untouched')
    const newRequest = await prepare(f.adapter)
    const inserted = await newRequest.execute(
      ops({ op: 'insertHorizontalRule', after: 'selection' }),
    )
    expect(inserted.result.isError).toBeUndefined()
  })

  it('rejects document and metadata changes even for selection-addressed operations', async () => {
    const f = fixture()
    select(f.editor, 'old')
    const session = await prepare(f.adapter)
    f.editor.commands.insertContentAt(f.editor.state.doc.content.size, 'external user change')
    const changed = f.editor.getJSON()
    await expect(
      session.execute(ops({ op: 'deleteBlocks', target: 'selection' })),
    ).rejects.toMatchObject({ kind: 'stale-document' })
    expect(f.editor.getJSON()).toEqual(changed)
    const metadata = fixture()
    const metadataSession = await prepare(metadata.adapter)
    metadata.frontmatter.write('title: external change')
    await expect(
      metadataSession.execute(ops({ op: 'setFrontmatter', yaml: 'title: overwrite' })),
    ).rejects.toMatchObject({ kind: 'stale-document' })
    expect(metadata.frontmatter.read()).toBe('title: external change')
  })

  it('invalidates the captured editor, document load and cancelled operations before editing', async () => {
    const replacement = fixture()
    const a = await prepare(replacement.adapter)
    replacement.replaceEditor(editorOf('another editor'))
    await expect(
      a.execute(ops({ op: 'deleteBlocks', target: { start: 0 } })),
    ).rejects.toMatchObject({ kind: 'stale-document' })
    const reload = fixture()
    const b = await prepare(reload.adapter)
    reload.replaceDocument()
    await expect(
      b.execute(ops({ op: 'deleteBlocks', target: { start: 0 } })),
    ).rejects.toMatchObject({ kind: 'stale-document' })
    const cancelled = fixture()
    const c = await prepare(cancelled.adapter)
    const controller = new AbortController()
    controller.abort()
    const before = cancelled.editor.getJSON()
    await expect(
      c.execute(ops({ op: 'deleteBlocks', target: { start: 0 } }), controller.signal),
    ).rejects.toMatchObject({ kind: 'cancelled' })
    expect(cancelled.editor.getJSON()).toEqual(before)
  })

  it('checks cancellation again after restoring the captured range and returns the visible caret', async () => {
    const f = fixture()
    select(f.editor, 'old')
    const state = await f.adapter.capture()
    f.editor.commands.setTextSelection(positionOf(f.editor, 'beta'))
    const visible = f.editor.state.selection
    const before = f.editor.getJSON()
    const session = await f.adapter.prepare({ runId: 'cancel-on-scope', mode: 'edit', state })
    const controller = new AbortController()
    const cancel = () => controller.abort()
    f.editor.on('selectionUpdate', cancel)
    try {
      await expect(
        session.execute(ops({ op: 'deleteBlocks', target: 'selection' }), controller.signal),
      ).rejects.toMatchObject({ kind: 'cancelled' })
      expect(f.editor.getJSON()).toEqual(before)
      expect(f.editor.state.selection.eq(visible)).toBe(true)
    } finally {
      f.editor.off('selectionUpdate', cancel)
    }
  })

  it('exposes no write tool for a reference or read-only document', async () => {
    const reference = fixture()
    const session = await prepare(reference.adapter, 'reference')
    expect(session.tools.map((tool) => tool.name)).not.toContain('apply_ops')
    await expect(
      session.execute(ops({ op: 'deleteBlocks', target: 'selection' })),
    ).rejects.toMatchObject({ kind: 'unknown-tool' })
    const readonly = fixture()
    readonly.editor.setEditable(false)
    const ro = await prepare(readonly.adapter)
    expect(ro.tools.map((tool) => tool.name)).not.toContain('apply_ops')
    const changed = fixture()
    const editing = await prepare(changed.adapter)
    changed.editor.setEditable(false)
    await expect(
      editing.execute(ops({ op: 'deleteBlocks', target: { start: 0 } })),
    ).rejects.toMatchObject({ kind: 'read-only' })
  })

  it('rejects a private frontmatter write without changing earlier batch operations', async () => {
    const f = fixture()
    const session = await prepare(f.adapter)
    const before = f.editor.getJSON()
    await expect(
      session.execute(
        ops(
          { op: 'replaceText', target: { start: 1 }, find: 'old', replace: 'changed' },
          { op: 'setFrontmatter', yaml: 'key: <span data-redaction="">private</span>' },
        ),
      ),
    ).rejects.toMatchObject({ kind: 'unsafe-content' })
    expect(f.editor.getJSON()).toEqual(before)
    expect(f.frontmatter.read()).toBe('')
  })
})
