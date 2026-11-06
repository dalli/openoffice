import { afterEach, describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import { nextNoteId, parseDocx, saveDocx } from '@genoffice/docx-engine'
import { odtToDocx } from '@genoffice/odt-engine'
import type { AgentToolCall, WorkspaceDocumentAdapter } from '@genoffice/agent-core'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { buildOdt } from '../../../packages/odt-engine/tests/helpers/build-odt'
import { hPara } from '../../../packages/odt-engine/tests/helpers/hancom-fragments'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { addCommentToSelection } from '../src/renderer/editor/comments'
import { blocksToPmDoc, pmDocToSavePlan, type PmNode } from '../src/renderer/editor/convert'
import { DocsReviewStore } from '../src/renderer/review-store'
import {
  addCommentAt,
  deleteComment,
  replyToComment,
  resolveComment,
  type ReviewContext,
} from '../src/renderer/review-actions'
import { createDocsSkill } from '../src/renderer/ai/docs-skill'
import { createDocsWorkspaceAdapter } from '../src/renderer/ai/workspace-adapter'
import type { AiCommentsAccess } from '../src/renderer/ai/tools'
import type { AiNotesAccess } from '../src/renderer/ai/note-ops'
import { beginOdtSession, endOdtSession } from '../src/renderer/odt-gate'

const editors: Editor[] = []
const adapters: WorkspaceDocumentAdapter[] = []
let sequence = 0
const tool = (name: string, input: Record<string, unknown> = {}): AgentToolCall => ({
  id: `review-${++sequence}`,
  name,
  input,
})
const NUM_IDS = { bullet: null, ordered: null }

function fixture() {
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: '<p>First paragraph</p><p>Second paragraph</p>',
  })
  editors.push(editor)
  const store = new DocsReviewStore()
  const ctx = {
    editor,
    get comments() {
      return store.getSnapshot().comments
    },
    setComments: store.setComments,
    setCommentsDirty: store.setCommentsDirty,
    dirtyRef: { current: false },
    setStatus: () => {},
  } as unknown as ReviewContext
  const comments: AiCommentsAccess = {
    list: () => store.getSnapshot().comments,
    add: (range, text, meta) =>
      addCommentAt(ctx, range, text, meta.author ?? 'Fixture AI', meta.initials),
    reply: (id, text) => replyToComment(ctx, id, text, 'Fixture AI'),
    resolve: (id) => {
      if (!comments.list().some((item) => item.id === id)) return false
      resolveComment(ctx, id, true)
      return true
    },
    remove: (id) => {
      if (!comments.list().some((item) => item.id === id)) return false
      deleteComment(ctx, id)
      return true
    },
  }
  const list: AiNotesAccess['list'] = (kind) =>
    store.getSnapshot()[kind === 'footnote' ? 'footnotes' : 'endnotes']
  const commit = (kind: 'footnote' | 'endnote', next: ReturnType<AiNotesAccess['list']>) => {
    ;(kind === 'footnote' ? store.setFootnotes : store.setEndnotes)(next)
    store.setNotesDirty(true)
  }
  const notes: AiNotesAccess = {
    list,
    add(kind, text) {
      const id = nextNoteId(list(kind))
      commit(kind, [...list(kind), { id, text }])
      return id
    },
    remove(kind, id) {
      if (!list(kind).some((item) => item.id === id)) return false
      commit(
        kind,
        list(kind).filter((item) => item.id !== id),
      )
      return true
    },
    replace(kind, id, next) {
      if (!list(kind).some((item) => item.id === id)) return false
      commit(
        kind,
        list(kind).map((item) => (item.id === id ? next : item)),
      )
      return true
    },
  }
  const deps = {
    getEditor: () => editor,
    getDocumentId: () => 'review-document',
    getNumIds: () => NUM_IDS,
    reviewStore: store,
    getComments: () => comments,
    getNotes: () => notes,
  }
  const adapter = createDocsWorkspaceAdapter(deps)
  adapters.push(adapter)
  const direct = createDocsSkill(
    () => editor,
    () => NUM_IDS,
    undefined,
    () => comments,
    undefined,
    () => false,
    undefined,
    undefined,
    undefined,
    () => notes,
  )
  return { editor, store, adapter, deps, direct }
}

async function prepare(adapter: WorkspaceDocumentAdapter, mode: 'edit' | 'reference' = 'edit') {
  return adapter.prepare({
    runId: `review-run-${++sequence}`,
    mode,
    state: await adapter.capture(),
  })
}

afterEach(() => {
  for (const adapter of adapters.splice(0)) adapter.dispose()
  for (const editor of editors.splice(0)) editor.destroy()
  endOdtSession()
  vi.useRealTimers()
})

describe('workspace review stores', () => {
  it('marks only changed review lists dirty when restoring after dirty flags were cleared', () => {
    const store = new DocsReviewStore()
    const before = store.capture()
    store.setComments([{ id: '1', author: 'AI', text: 'Added' }])
    store.setCommentsDirty(false)
    const revision = store.revision
    store.restore(before)
    expect(store.getSnapshot().comments).toEqual([])
    expect(store.getSnapshot().commentsDirty).toBe(true)
    expect(store.getSnapshot().notesDirty).toBe(false)
    expect(store.getSnapshot().footnotes).toBe(before.footnotes)
    expect(store.getSnapshot().endnotes).toBe(before.endnotes)
    expect(store.revision).toBe(revision + 1)
  })

  it('matches the original comment and note executors across a batch without React renders', async () => {
    vi.useFakeTimers({ toFake: ['Date'] })
    vi.setSystemTime(new Date('2026-11-05T00:00:00Z'))
    const original = fixture()
    const adapted = fixture()
    original.direct.buildContext?.()
    const session = await prepare(adapted.adapter)
    for (const call of [
      tool('add_comment', { blockIndex: 0, text: 'First', comment: 'Please clarify' }),
      tool('reply_comment', { parentId: '1', text: 'First reply' }),
      tool('reply_comment', { parentId: '1', text: 'Second reply' }),
      tool('read_comments'),
      tool('resolve_comment', { id: '1' }),
      tool('insert_footnote', { blockIndex: 0, text: 'First footnote' }),
      tool('insert_footnote', { blockIndex: 0, text: 'Second footnote' }),
      tool('insert_endnote', { blockIndex: 1, text: 'First endnote' }),
      tool('edit_note', { kind: 'footnote', id: '1', find: 'First', replace: 'Edited' }),
      tool('read_notes'),
      tool('delete_note', { kind: 'footnote', id: '2' }),
      tool('delete_comment', { id: '1', withReplies: true }),
    ]) {
      const beforeRevision = adapted.store.revision
      const expected = await original.direct.executeTool(call)
      const actual = await session.execute(call)
      expect(actual.result, call.name).toEqual({
        ...expected,
        mutated: expected.mutated || beforeRevision !== adapted.store.revision,
      })
      expect(actual.result.isError, call.name).not.toBe(true)
      expect(adapted.editor.getJSON(), call.name).toEqual(original.editor.getJSON())
      expect(adapted.store.getSnapshot(), call.name).toEqual(original.store.getSnapshot())
    }
  })

  it('undoes store-only resolution and note edits without undoing earlier manual content', async () => {
    const { editor, store, adapter, direct } = fixture()
    direct.buildContext?.()
    await direct.executeTool(tool('add_comment', { blockIndex: 0, comment: 'Manual comment' }))
    await direct.executeTool(tool('insert_footnote', { blockIndex: 0, text: 'Manual footnote' }))
    const before = editor.getJSON()
    const beforeNode = editor.state.doc
    const review = store.capture()
    const pinned = await adapter.capture()
    const session = await adapter.prepare({ runId: 'store-only', mode: 'edit', state: pinned })
    const resolved = await session.execute(tool('resolve_comment', { id: '1' }))
    expect(resolved.result.mutated).toBe(true)
    expect(resolved.state.revision).not.toBe(pinned.revision)
    const edited = await session.execute(
      tool('edit_note', { kind: 'footnote', id: '1', find: 'Manual', replace: 'AI' }),
    )
    expect(edited.result.mutated).toBe(true)
    expect(editor.getJSON()).toEqual(before)
    expect(store.getSnapshot().comments[0].done).toBe(true)
    expect(store.getSnapshot().footnotes[0].text).toBe('AI footnote')
    session.release()
    await adapter.rollback!(session.rollbackId!, session.state)
    expect(editor.getJSON()).toEqual(before)
    expect(editor.state.doc).toBe(beforeNode)
    expect(store.getSnapshot()).toEqual(review)
    expect((await adapter.refresh(pinned)).selectionId).toBe(pinned.selectionId)
  })

  it('restores comment anchors and both note lists together with body edits', async () => {
    const { editor, store, adapter } = fixture()
    const before = editor.getJSON()
    const session = await prepare(adapter)
    await session.execute(tool('add_comment', { blockIndex: 0, comment: 'AI comment' }))
    await session.execute(tool('insert_footnote', { blockIndex: 0, text: 'AI footnote' }))
    await session.execute(tool('insert_endnote', { blockIndex: 1, text: 'AI endnote' }))
    await session.execute(tool('insert_content', { html: '<p>AI body</p>' }))
    session.release()
    await adapter.rollback!(session.rollbackId!, session.state)
    expect(editor.getJSON()).toEqual(before)
    expect(store.getSnapshot()).toEqual({
      comments: [],
      footnotes: [],
      endnotes: [],
      commentsDirty: true,
      notesDirty: true,
    })
  })

  it.each(['comments', 'footnotes', 'endnotes'] as const)(
    'invalidates queued captures and rollback after manual %s changes',
    async (key) => {
      const { editor, store, adapter } = fixture()
      const pinned = await adapter.capture()
      const session = await prepare(adapter)
      await session.execute(tool('insert_content', { html: '<p>AI body</p>' }))
      session.release()
      const unchanged = editor.getJSON()
      if (key === 'comments')
        store.setComments([{ id: '7', author: 'Manual', text: 'Manual comment' }])
      else
        (key === 'footnotes' ? store.setFootnotes : store.setEndnotes)([
          { id: '7', text: 'Manual note' },
        ])
      await expect(
        adapter.prepare({ runId: 'queued', mode: 'edit', state: pinned }),
      ).rejects.toMatchObject({ kind: 'stale-document' })
      await expect(adapter.rollback!(session.rollbackId!, session.state)).rejects.toMatchObject({
        kind: 'stale-document',
      })
      expect(editor.getJSON()).toEqual(unchanged)
      expect(store.getSnapshot()[key][0].text).toContain('Manual')
    },
  )

  it('publishes only content invalidations and removes observers on disposal', async () => {
    const { store, adapter } = fixture()
    const notify = vi.fn()
    adapter.subscribeChanges(notify)
    const state = await adapter.capture()
    store.setCommentsDirty(true)
    store.setNotesDirty(true)
    expect((await adapter.capture()).revision).toBe(state.revision)
    expect(notify).not.toHaveBeenCalled()
    store.setFootnotes([{ id: '1', text: 'Manual' }])
    expect(notify).toHaveBeenCalledOnce()
    adapter.dispose()
    store.setFootnotes([])
    expect(notify).toHaveBeenCalledOnce()
  })

  it.each(['reference', 'read-only'] as const)(
    'keeps review writes unavailable for %s documents',
    async (mode) => {
      const { editor, adapter } = fixture()
      if (mode === 'read-only') editor.setEditable(false)
      const session = await prepare(adapter, mode === 'reference' ? mode : 'edit')
      expect(session.tools.map((item) => item.name)).toEqual(
        expect.arrayContaining(['read_comments', 'read_notes']),
      )
      expect(session.tools.map((item) => item.name)).not.toContain('resolve_comment')
      await expect(
        session.execute(tool('insert_footnote', { blockIndex: 0, text: 'Forbidden' })),
      ).rejects.toMatchObject({ kind: 'unknown-tool' })
    },
  )

  it('rejects a review write if the document becomes read-only after preparation', async () => {
    const { editor, store, adapter } = fixture()
    const session = await prepare(adapter)
    editor.setEditable(false)
    await expect(
      session.execute(tool('insert_footnote', { blockIndex: 0, text: 'Late note' })),
    ).rejects.toMatchObject({ kind: 'read-only' })
    expect(store.getSnapshot().footnotes).toEqual([])
  })

  it('retains the original ODT limits with review stores connected', async () => {
    const Native = globalThis.TextEncoder
    class RealmEncoder extends Native {
      encode(input?: string): Uint8Array<ArrayBuffer> {
        return new Uint8Array(super.encode(input))
      }
    }
    globalThis.TextEncoder = RealmEncoder
    let converted: Awaited<ReturnType<typeof odtToDocx>>
    try {
      converted = await odtToDocx(await buildOdt({ body: hPara('ODT review fixture') }))
    } finally {
      globalThis.TextEncoder = Native
    }
    beginOdtSession({ bridge: converted.bridge }, await parseDocx(converted.docxBytes))
    const { store, adapter } = fixture()
    const session = await prepare(adapter)
    const names = session.tools.map((entry) => entry.name)
    expect(names).toEqual(
      expect.arrayContaining(['read_comments', 'read_notes', 'insert_footnote']),
    )
    for (const name of ['add_comment', 'reply_comment', 'resolve_comment']) {
      expect(names).not.toContain(name)
      await expect(session.execute(tool(name, {}))).rejects.toMatchObject({ kind: 'unknown-tool' })
    }
    expect(store.getSnapshot().comments).toEqual([])
  })

  it('rechecks review edits after the asynchronous first-write hook', async () => {
    const { store, adapter, deps } = fixture()
    const guarded = createDocsWorkspaceAdapter({
      ...deps,
      beforeFirstWrite: async () => {
        store.setComments([{ id: '9', author: 'Manual', text: 'Changed during preparation' }])
      },
    })
    adapters.push(guarded)
    const session = await prepare(guarded)
    await expect(
      session.execute(tool('insert_footnote', { blockIndex: 0, text: 'Late' })),
    ).rejects.toMatchObject({ kind: 'stale-document' })
    expect(store.getSnapshot().footnotes).toEqual([])
    expect((await adapter.capture()).revision).not.toBe('0')
  })

  it('projects redacted comment anchors and keeps references limited to the requested body', async () => {
    const { editor, store, adapter } = fixture()
    editor.commands.setTextSelection({ from: 1, to: 6 })
    editor.commands.setRedaction('private anchor')
    addCommentToSelection(editor, '7')
    store.setComments([{ id: '7', author: 'Manual', text: 'Separate comment' }])
    store.setFootnotes([{ id: '1', text: 'Separate note' }])
    const session = await prepare(adapter)
    expect(session.context).not.toContain('First')
    expect(session.context).toContain('{{private anchor}}')
    expect(session.context).toContain('Separate comment')
    const read = await session.execute(tool('read_comments'))
    expect(read.result.output).not.toContain('First')
    expect(read.result.output).toContain('{{private anchor}}')
    session.release()
    const reference = await adapter.read(await adapter.capture(), { scope: 'document' })
    expect(reference.text).not.toContain('Separate comment')
    expect(reference.text).not.toContain('Separate note')
  })

  it('saves the same review parts as the original executor and restores the before-run content', async () => {
    const source = await buildDocx({ bodyXml: '<w:p><w:r><w:t>First paragraph</w:t></w:r></w:p>' })
    const parsed = await parseDocx(source)
    const original = fixture()
    const adapted = fixture()
    for (const item of [original, adapted])
      item.editor.commands.setContent(blocksToPmDoc(parsed.blocks) as never)
    const bytes = async (item: ReturnType<typeof fixture>) => {
      const state = item.store.getSnapshot()
      const plan = pmDocToSavePlan(item.editor.getJSON() as PmNode, parsed.blocks)
      return saveDocx(parsed, plan.saveBlocks, {
        comments: state.commentsDirty ? state.comments : undefined,
        footnotes: state.notesDirty ? state.footnotes : undefined,
        endnotes: state.notesDirty ? state.endnotes : undefined,
      })
    }
    vi.useFakeTimers({ toFake: ['Date'] })
    vi.setSystemTime(new Date('2026-11-05T00:00:00Z'))
    original.direct.buildContext?.()
    const before = adapted.editor.getJSON()
    const session = await prepare(adapted.adapter)
    for (const call of [
      tool('add_comment', { blockIndex: 0, comment: 'Saved comment' }),
      tool('insert_footnote', { blockIndex: 0, text: 'Saved footnote' }),
      tool('insert_endnote', { blockIndex: 0, text: 'Saved endnote' }),
    ]) {
      await original.direct.executeTool(call)
      await session.execute(call)
    }
    expect(await bytes(adapted)).toEqual(await bytes(original))
    const saved = await parseDocx(await bytes(adapted))
    expect(saved.comments[0].text).toBe('Saved comment')
    expect(saved.footnotes[0].text).toBe('Saved footnote')
    expect(saved.endnotes[0].text).toBe('Saved endnote')
    session.release()
    await adapted.adapter.rollback!(session.rollbackId!, session.state)
    const restored = await parseDocx(await bytes(adapted))
    expect(restored.comments).toEqual([])
    expect(restored.footnotes).toEqual([])
    expect(restored.endnotes).toEqual([])
    expect(adapted.editor.getJSON()).toEqual(before)
    expect(parsed.internal.originalBytes).toEqual(source)
  })
})
