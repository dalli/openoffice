import { afterEach, describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import { TextSelection } from '@tiptap/pm/state'
import { parseDocx } from '@genoffice/docx-engine'
import { odtToDocx } from '@genoffice/odt-engine'
import type { AgentToolCall, WorkspaceDocumentAdapter } from '@genoffice/agent-core'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { buildOdt } from '../../../packages/odt-engine/tests/helpers/build-odt'
import { hPara } from '../../../packages/odt-engine/tests/helpers/hancom-fragments'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { beginOdtSession, endOdtSession } from '../src/renderer/odt-gate'
import { createDocsSkill } from '../src/renderer/ai/docs-skill'
import {
  createDocsWorkspaceAdapter,
  type DocsWorkspaceAdapterDeps,
} from '../src/renderer/ai/workspace-adapter'

const NUM_IDS = { bullet: null, ordered: null }
const editors: Editor[] = []
const adapters: WorkspaceDocumentAdapter[] = []
const SECRET = 'hidden-fixture-phone'
let callSequence = 0

function editor(texts = ['First paragraph', 'Second paragraph']): Editor {
  const instance = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: {
      type: 'doc',
      content: texts.map((text) => ({
        type: 'docParagraph',
        content: [{ type: 'text', text }],
      })),
    },
  })
  editors.push(instance)
  return instance
}

function connect(initial: Editor, extra: Partial<DocsWorkspaceAdapterDeps> = {}) {
  let current = initial
  let documentId = 'loaded-document-1'
  const adapter = createDocsWorkspaceAdapter({
    getEditor: () => current,
    getDocumentId: () => documentId,
    getNumIds: () => NUM_IDS,
    ...extra,
  })
  adapters.push(adapter)
  return {
    adapter,
    replace(next: Editor) {
      current = next
    },
    reload() {
      documentId += '-reloaded'
    },
  }
}

function select(instance: Editor, needle: string): void {
  let from = -1
  instance.state.doc.descendants((node, pos) => {
    if (from < 0 && node.isText && node.text?.includes(needle))
      from = pos + node.text.indexOf(needle)
  })
  if (from < 0) throw new Error('selection fixture is missing')
  instance.view.dispatch(
    instance.state.tr.setSelection(
      TextSelection.create(instance.state.doc, from, from + needle.length),
    ),
  )
}

const tool = (name: string, input: Record<string, unknown>): AgentToolCall => ({
  id: `call-${++callSequence}`,
  name,
  input,
})
async function prepare(adapter: WorkspaceDocumentAdapter, mode: 'edit' | 'reference' = 'edit') {
  const state = await adapter.capture()
  return adapter.prepare({ runId: `run-${++callSequence}`, mode, state })
}

afterEach(() => {
  for (const adapter of adapters.splice(0)) adapter.dispose()
  for (const instance of editors.splice(0)) instance.destroy()
  endOdtSession()
  vi.restoreAllMocks()
})

describe('DOCX workspace adapter on the real Tiptap editor', () => {
  it('reads current unsaved content and projects known redactions without touching source bytes', async () => {
    const source = await buildDocx({ bodyXml: `<w:p><w:r><w:t>Call ${SECRET}</w:t></w:r></w:p>` })
    const originalBytes = source.slice()
    const parsed = await parseDocx(source)
    const instance = editor()
    instance.commands.setContent({
      type: 'doc',
      content: blocksToPmDoc(parsed.blocks).content as never[],
    })
    select(instance, SECRET)
    instance.commands.setRedaction('client phone')
    instance.view.dispatch(
      instance.state.tr.insertText(' UNSAVED', instance.state.doc.content.size - 1),
    )
    const { adapter } = connect(instance)
    const state = await adapter.capture()
    const read = await adapter.read(state, { scope: 'document' })
    expect(read.text).toBe('Call {{client phone}} UNSAVED')
    expect(read.text).not.toContain(SECRET)
    expect(read.masking).toBe('applied')
    expect(instance.state.doc.textContent).toContain(SECRET)
    expect(source).toEqual(originalBytes)
    const session = await adapter.prepare({ runId: 'redacted', mode: 'reference', state })
    expect(session.context).not.toContain(SECRET)
    const result = await session.execute(
      tool('read_blocks', { startBlockIndex: 0, endBlockIndex: 0 }),
    )
    expect(result.result.output).toContain('{{client phone}}')
    expect(result.result.output).not.toContain(SECRET)
  })

  it('retains the captured selection through panel focus and a different live caret', async () => {
    const direct = editor()
    const adapted = editor()
    for (const instance of [direct, adapted]) select(instance, 'First')
    const original = createDocsSkill(
      () => direct,
      () => NUM_IDS,
    )
    original.buildContext?.()
    const { adapter } = connect(adapted)
    const captured = await adapter.capture()
    for (const instance of [direct, adapted]) select(instance, 'Second')
    const liveSelection = adapted.state.selection
    const session = await adapter.prepare({
      runId: 'frozen-selection',
      mode: 'edit',
      state: captured,
    })
    expect(adapted.state.selection).toBe(liveSelection)
    expect(session.context).toContain('<sel>First</sel>')
    const read = await adapter.read(captured, { scope: 'selection' })
    expect(read.text).toBe('First')
    const call = tool('replace_selection', { html: '<strong>Updated</strong>' })
    const expected = await original.executeTool(call)
    const actual = await session.execute(call)
    expect(actual.result).toEqual(expected)
    expect(adapted.getJSON()).toEqual(direct.getJSON())
    expect(adapted.state.doc.child(1).textContent).toBe('Second paragraph')
  })

  it('reads a current table with its withheld cell text projected', async () => {
    const source = await buildDocx({
      bodyXml:
        '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="2000"/></w:tblGrid>' +
        `<w:tr><w:tc><w:p><w:r><w:t>${SECRET}</w:t></w:r></w:p></w:tc>` +
        '<w:tc><w:p><w:r><w:t>Other cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>',
    })
    const parsed = await parseDocx(source)
    const instance = editor()
    instance.commands.setContent({
      type: 'doc',
      content: blocksToPmDoc(parsed.blocks).content as never[],
    })
    select(instance, SECRET)
    instance.commands.setRedaction('private cell')
    const { adapter } = connect(instance)
    const read = await adapter.read(await adapter.capture(), { scope: 'document' })
    expect(read.text).toContain('{{private cell}}')
    expect(read.text).toContain('Other cell')
    expect(read.text).not.toContain(SECRET)
    const session = await prepare(adapter, 'reference')
    const result = await session.execute(
      tool('read_blocks', { startBlockIndex: 0, endBlockIndex: 0 }),
    )
    expect(result.result.output).toContain('<table>')
    expect(result.result.output).toContain('{{private cell}}')
    expect(result.result.output).not.toContain(SECRET)
  })

  it('matches the original skill for reads, formatting, replacements and insertion', async () => {
    const direct = editor()
    const adapted = editor()
    const original = createDocsSkill(
      () => direct,
      () => NUM_IDS,
    )
    original.buildContext?.()
    const session = await prepare(connect(adapted).adapter)
    for (const call of [
      tool('read_blocks', { startBlockIndex: 0, endBlockIndex: 1 }),
      tool('apply_ops', { ops: [{ op: 'setFont', target: { blockIndexes: [0] }, bold: true }] }),
      tool('replace_blocks', {
        startBlockIndex: 1,
        endBlockIndex: 1,
        html: '<p>Rewritten body</p>',
      }),
      tool('insert_content', { html: '<p>Added paragraph</p>', afterBlockIndex: 1 }),
    ]) {
      const expected = await original.executeTool(call)
      const actual = await session.execute(call)
      expect(actual.result, call.name).toEqual(expected)
      expect(adapted.getJSON(), call.name).toEqual(direct.getJSON())
    }
    expect(await session.verifyResponse('Updated the document.', [])).toBeNull()
  })

  it('keeps the intended range after its own formatting edit and another live caret move', async () => {
    const instance = editor()
    select(instance, 'First')
    const session = await prepare(connect(instance).adapter)
    await session.execute(
      tool('apply_ops', {
        ops: [{ op: 'setFont', target: { scope: 'selection' }, bold: true }],
      }),
    )
    select(instance, 'Second')
    await session.execute(tool('replace_selection', { html: 'Correct' }))
    expect(instance.state.doc.child(0).textContent).toBe('Correct paragraph')
    expect(instance.state.doc.child(1).textContent).toBe('Second paragraph')
    select(instance, 'Second')
    await session.execute(tool('replace_selection', { html: 'Again' }))
    expect(instance.state.doc.child(0).textContent).toBe('Again paragraph')
    expect(instance.state.doc.child(1).textContent).toBe('Second paragraph')
  })

  it('keeps a multi-op change in the same undo unit as the original executor', async () => {
    const direct = editor()
    const adapted = editor()
    const before = adapted.getJSON()
    const original = createDocsSkill(
      () => direct,
      () => NUM_IDS,
    )
    original.buildContext?.()
    const session = await prepare(connect(adapted).adapter)
    const call = tool('apply_ops', {
      ops: [
        { op: 'setFont', target: { blockIndexes: [0] }, bold: true },
        { op: 'setFont', target: { blockIndexes: [1] }, italic: true },
      ],
    })
    await original.executeTool(call)
    await session.execute(call)
    expect(adapted.getJSON()).toEqual(direct.getJSON())
    expect(adapted.commands.undo()).toBe(true)
    expect(direct.commands.undo()).toBe(true)
    expect(adapted.getJSON()).toEqual(before)
    expect(adapted.getJSON()).toEqual(direct.getJSON())
  })

  it('preserves AI change tracking and the original revision markers', async () => {
    const direct = editor()
    const adapted = editor()
    const track = () => ({ author: 'Workspace fixture AI' })
    const original = createDocsSkill(
      () => direct,
      () => NUM_IDS,
      track,
    )
    original.buildContext?.()
    const session = await prepare(connect(adapted, { getTrack: track }).adapter)
    const call = tool('replace_blocks', {
      startBlockIndex: 0,
      endBlockIndex: 0,
      html: '<p>AI replacement</p>',
    })
    const expected = await original.executeTool(call)
    const actual = await session.execute(call)
    const withoutDates = (value: unknown) =>
      JSON.stringify(value).replace(/"date":"[^"]*"/g, '"date":"<same-clock>"')
    expect(actual.result).toEqual(expected)
    expect(withoutDates(adapted.getJSON())).toEqual(withoutDates(direct.getJSON()))
    const marks: string[] = []
    adapted.state.doc.descendants((node) => {
      for (const mark of node.marks)
        if (mark.attrs.author === 'Workspace fixture AI') marks.push(mark.type.name)
    })
    expect(marks).toContain('ins')
    expect(marks).toContain('del')
    const revisions = await session.execute(tool('read_revisions', {}))
    expect(revisions.result.output).toContain('Workspace fixture AI')
  })

  it('rejects a captured document after a user edit, including edit followed by undo', async () => {
    const instance = editor()
    const { adapter } = connect(instance)
    const captured = await adapter.capture()
    const before = instance.getJSON()
    instance.view.dispatch(instance.state.tr.insertText('manual ', 1))
    expect(instance.commands.undo()).toBe(true)
    expect(instance.getJSON()).toEqual(before)
    await expect(
      adapter.prepare({ runId: 'old', mode: 'edit', state: captured }),
    ).rejects.toMatchObject({ kind: 'stale-document' })
  })

  it('rejects manual changes made after preparation and preserves those changes', async () => {
    const instance = editor()
    const session = await prepare(connect(instance).adapter)
    instance.view.dispatch(instance.state.tr.insertText('manual ', 1))
    const manual = instance.getJSON()
    await expect(
      session.execute(
        tool('replace_blocks', {
          startBlockIndex: 0,
          endBlockIndex: 0,
          html: '<p>Wrong result</p>',
        }),
      ),
    ).rejects.toMatchObject({ kind: 'stale-document' })
    expect(instance.getJSON()).toEqual(manual)
  })

  it('rejects a forged selection handle rather than using the live selection', async () => {
    const instance = editor()
    const { adapter } = connect(instance)
    const captured = await adapter.capture()
    await expect(
      adapter.prepare({
        runId: 'bad-selection',
        mode: 'edit',
        state: { ...captured, selectionId: 'unknown-selection' },
      }),
    ).rejects.toMatchObject({ kind: 'stale-selection' })
  })

  it.each(['reload', 'replace'] as const)(
    'rejects %s even with the same visible document text',
    async (action) => {
      const instance = editor()
      const host = connect(instance)
      const session = await prepare(host.adapter)
      if (action === 'reload') host.reload()
      else host.replace(editor())
      await expect(
        session.execute(tool('insert_content', { html: '<p>Late result</p>' })),
      ).rejects.toMatchObject({ kind: 'stale-document' })
      expect(instance.state.doc.textContent).toBe('First paragraphSecond paragraph')
    },
  )

  it('checks cancellation directly before applying a document tool', async () => {
    const instance = editor()
    const session = await prepare(connect(instance).adapter)
    const before = instance.getJSON()
    const controller = new AbortController()
    controller.abort()
    await expect(
      session.execute(tool('insert_content', { html: '<p>Cancelled</p>' }), controller.signal),
    ).rejects.toMatchObject({ kind: 'cancelled' })
    expect(instance.getJSON()).toEqual(before)
  })

  it('refuses a late tool call after the original editor has been destroyed', async () => {
    const instance = editor()
    const session = await prepare(connect(instance).adapter)
    instance.destroy()
    await expect(
      session.execute(tool('insert_content', { html: '<p>Late result</p>' })),
    ).rejects.toMatchObject({ kind: 'not-ready' })
  })

  it('only exposes reading tools to a reference document', async () => {
    const instance = editor()
    const session = await prepare(connect(instance).adapter, 'reference')
    expect(session.tools.map((entry) => entry.name).sort()).toEqual([
      'get_document_context',
      'read_blocks',
      'read_revisions',
    ])
    await expect(
      session.execute(tool('insert_content', { html: '<p>Not allowed</p>' })),
    ).rejects.toMatchObject({ kind: 'unknown-tool' })
    expect(instance.state.doc.textContent).toBe('First paragraphSecond paragraph')
  })

  it('honors read-only state both before and after preparation', async () => {
    const instance = editor()
    const { adapter } = connect(instance)
    instance.setEditable(false)
    const readOnly = await prepare(adapter)
    expect(
      readOnly.tools.every((entry) =>
        ['get_document_context', 'read_blocks', 'read_revisions'].includes(entry.name),
      ),
    ).toBe(true)
    readOnly.release()
    instance.setEditable(true)
    const session = await prepare(adapter)
    instance.setEditable(false)
    await expect(
      session.execute(tool('insert_content', { html: '<p>Locked</p>' })),
    ).rejects.toMatchObject({ kind: 'read-only' })
  })

  it('never exposes network, media, file creation, saving or side-store tools', async () => {
    const session = await prepare(connect(editor()).adapter)
    const names = session.tools.map((entry) => entry.name)
    for (const name of [
      'web_search',
      'read_url',
      'image_search',
      'insert_image',
      'generate_image',
      'analyze_media',
      'create_document',
      'write_document',
      'save_document',
      'read_comments',
      'set_header_footer',
      'read_notes',
    ]) {
      expect(names).not.toContain(name)
      await expect(session.execute(tool(name, {}))).rejects.toMatchObject({ kind: 'unknown-tool' })
    }
  })

  it.each([null, '', '   ', '{{ambiguous}}', 'x'.repeat(41)])(
    'fails closed for a malformed redaction label %j',
    async (label) => {
      const instance = editor([`Before ${SECRET} after`])
      select(instance, SECRET)
      const { from, to } = instance.state.selection
      instance.view.dispatch(
        instance.state.tr.addMark(from, to, instance.schema.marks.redaction.create({ label })),
      )
      const { adapter } = connect(instance)
      await expect(adapter.capture()).rejects.toMatchObject({ kind: 'unsafe-content' })
      expect(instance.state.doc.textContent).toContain(SECRET)
    },
  )

  it('refuses a redacted whole-block deletion before the legacy context reader can expose a preview', async () => {
    const instance = editor([SECRET])
    select(instance, SECRET)
    instance.commands.setRedaction('private deletion')
    const { from, to } = instance.state.selection
    instance.view.dispatch(
      instance.state.tr.addMark(from, to, instance.schema.marks.del.create({ author: 'Fixture' })),
    )
    await expect(connect(instance).adapter.capture()).rejects.toMatchObject({
      kind: 'unsafe-content',
    })
  })

  it('refuses tracked writes to a redacted document before applying an unsafe deletion preview', async () => {
    const instance = editor([SECRET])
    select(instance, SECRET)
    instance.commands.setRedaction('private replacement')
    const session = await prepare(
      connect(instance, { getTrack: () => ({ author: 'Fixture AI' }) }).adapter,
    )
    const before = instance.getJSON()
    await expect(
      session.execute(
        tool('replace_blocks', {
          startBlockIndex: 0,
          endBlockIndex: 0,
          html: '<p>Replacement</p>',
        }),
      ),
    ).rejects.toMatchObject({ kind: 'unsafe-content' })
    expect(instance.getJSON()).toEqual(before)
  })

  it('reports a clipped read and refuses a selection read from a caret', async () => {
    const instance = editor()
    const { adapter } = connect(instance)
    const captured = await adapter.capture()
    const read = await adapter.read(captured, { scope: 'document', maxChars: 5 })
    expect(read.text).toBe('First')
    expect(read.truncated).toBe(true)
    await expect(adapter.read(captured, { scope: 'selection' })).rejects.toMatchObject({
      kind: 'unsupported-read',
    })
  })

  it('uses the original first-write hook once and rechecks edits made while it awaited', async () => {
    const instance = editor()
    const beforeFirstWrite = vi.fn<NonNullable<DocsWorkspaceAdapterDeps['beforeFirstWrite']>>(
      async () => {
        await Promise.resolve()
        instance.view.dispatch(instance.state.tr.insertText('manual ', 1))
      },
    )
    const session = await prepare(connect(instance, { beforeFirstWrite }).adapter)
    await expect(
      session.execute(tool('insert_content', { html: '<p>Late result</p>' })),
    ).rejects.toMatchObject({ kind: 'stale-document' })
    expect(beforeFirstWrite).toHaveBeenCalledOnce()
    expect(beforeFirstWrite.mock.calls[0][0]).toBe(instance)
    expect(instance.state.doc.textContent).not.toContain('Late result')
  })

  it('captures the original undo hook only once across several writes', async () => {
    const instance = editor()
    const beforeFirstWrite = vi.fn()
    const session = await prepare(connect(instance, { beforeFirstWrite }).adapter)
    await session.execute(
      tool('apply_ops', { ops: [{ op: 'setFont', target: { blockIndexes: [0] }, bold: true }] }),
    )
    await session.execute(tool('insert_content', { html: '<p>Added</p>', afterBlockIndex: 1 }))
    expect(beforeFirstWrite).toHaveBeenCalledOnce()
    expect(beforeFirstWrite).toHaveBeenCalledWith(instance, expect.any(AbortSignal))
  })

  it('keeps ODT tool filtering and nested operation refusals from the original skill', async () => {
    // JSZip needs byte arrays from the same realm as this jsdom test.
    const Native = globalThis.TextEncoder
    class RealmEncoder extends Native {
      encode(input?: string): Uint8Array<ArrayBuffer> {
        return new Uint8Array(super.encode(input))
      }
    }
    globalThis.TextEncoder = RealmEncoder
    let converted: Awaited<ReturnType<typeof odtToDocx>>
    try {
      converted = await odtToDocx(await buildOdt({ body: hPara('ODT fixture') }))
    } finally {
      globalThis.TextEncoder = Native
    }
    const parsed = await parseDocx(converted.docxBytes)
    beginOdtSession({ bridge: converted.bridge }, parsed)
    const instance = editor()
    instance.commands.setContent({
      type: 'doc',
      content: blocksToPmDoc(parsed.blocks).content as never[],
    })
    const session = await prepare(connect(instance).adapter)
    expect(session.systemPrompt).toContain('ODT')
    for (const name of ['insert_chart', 'edit_chart', 'accept_changes', 'reject_changes']) {
      expect(session.tools.some((entry) => entry.name === name)).toBe(false)
    }
    const before = instance.getJSON()
    const call = tool('apply_ops', { ops: [{ op: 'insertToc' }] })
    const original = createDocsSkill(
      () => instance,
      () => NUM_IDS,
    )
    original.buildContext?.()
    const expected = await original.executeTool(call)
    const actual = await session.execute(call)
    expect(actual.result).toEqual(expected)
    expect(actual.result.isError).toBe(true)
    expect(actual.result.output).toContain('not available in this ODT document')
    expect(instance.getJSON()).toEqual(before)
  })

  it('releases the execution session and removes its transaction listener on dispose', async () => {
    const instance = editor()
    const { adapter } = connect(instance)
    const session = await prepare(adapter)
    session.release()
    await expect(
      session.execute(tool('insert_content', { html: '<p>Late</p>' })),
    ).rejects.toMatchObject({ kind: 'released' })
    const off = vi.spyOn(instance, 'off')
    adapter.dispose()
    expect(off).toHaveBeenCalledWith('transaction', expect.any(Function))
    await expect(adapter.capture()).rejects.toMatchObject({ kind: 'disposed' })
  })
})
