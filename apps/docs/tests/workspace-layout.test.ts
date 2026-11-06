import { afterEach, describe, expect, it, vi } from 'vitest'
import { Editor } from '@tiptap/core'
import { odtToDocx } from '@genoffice/odt-engine'
import { buildOdt } from '../../../packages/odt-engine/tests/helpers/build-odt'
import { hPara } from '../../../packages/odt-engine/tests/helpers/hancom-fragments'
import { beginOdtSession, endOdtSession } from '../src/renderer/odt-gate'
import { parseDocx, readSections, readSectionSettings } from '@genoffice/docx-engine'
import type { AgentToolCall, WorkspaceDocumentAdapter } from '@genoffice/agent-core'
import { buildDocx } from '../../../packages/docx-engine/tests/helpers/build-docx'
import { editorExtensions } from '../src/renderer/editor/extensions'
import { blocksToPmDoc } from '../src/renderer/editor/convert'
import { DocsLayoutStore } from '../src/renderer/layout-store'
import { DocsReviewStore } from '../src/renderer/review-store'
import { buildDocBytes, type FileActionContext } from '../src/renderer/file-actions'
import { createDocsLayoutAccess } from '../src/renderer/ai/layout-access'
import { createDocsWorkspaceAdapter } from '../src/renderer/ai/workspace-adapter'
import { createDocsSkill } from '../src/renderer/ai/docs-skill'
import { executeTool } from '../src/renderer/ai/tools'
import type { DocState } from '../src/renderer/doc-state'

const adapters: WorkspaceDocumentAdapter[] = []
const editors: Editor[] = []
let sequence = 0
const call = (name: string, input: Record<string, unknown> = {}): AgentToolCall => ({
  id: `layout-${++sequence}`,
  name,
  input,
})
const NUM_IDS = { bullet: null, ordered: null }

async function fixture(source?: Uint8Array) {
  const original =
    source ??
    (await buildDocx({
      bodyXml:
        '<w:p><w:r><w:t>First paragraph</w:t></w:r></w:p><w:p><w:r><w:t>Second paragraph</w:t></w:r></w:p>',
    }))
  const parsed = await parseDocx(original)
  let doc: DocState = {
    parsed,
    filePath: '/fixture.docx',
    fileName: 'fixture.docx',
    hash: 'fixture',
    format: 'docx',
  }
  const editor = new Editor({
    element: document.createElement('div'),
    extensions: editorExtensions,
    content: blocksToPmDoc(parsed.blocks, readSections(parsed)) as never,
  })
  editors.push(editor)
  const store = new DocsLayoutStore()
  store.setSection(readSectionSettings(parsed))
  store.setSections(readSections(parsed))
  const review = new DocsReviewStore()
  const context = () =>
    ({
      ...store.getSnapshot(),
      ...review.getSnapshot(),
      editor,
      doc,
      inkAnnotations: [],
      inksDirty: false,
    }) as unknown as FileActionContext
  const insertBreak = vi.fn()
  let locked = false
  const access = createDocsLayoutAccess({
    store,
    getContext: context,
    isLocked: () => locked,
    getDeletedBreaks: () => undefined,
    insertBreak,
  })
  const deps = {
    getEditor: () => editor,
    getDocumentId: () => 'layout-fixture',
    getNumIds: () => NUM_IDS,
    layoutStore: store,
    getLayoutBaseline: () => doc.parsed,
    getHf: () => access.hf,
    getPageSetup: () => access.pageSetup(true),
    getExtras: () => access.extras,
    getTemplateDocument: () => ({
      fileName: doc.format === 'odt' ? 'fixture.odt' : 'fixture.docx',
      title: 'fixture',
      bytes: async () => (await buildDocBytes(context()))!,
      redacted: () => false,
    }),
  }
  const adapter = createDocsWorkspaceAdapter(deps)
  adapters.push(adapter)
  const direct = createDocsSkill(
    () => editor,
    () => NUM_IDS,
    undefined,
    undefined,
    () => access.hf,
    () => false,
    undefined,
    () => access.pageSetup(true),
    () => access.extras,
  )
  return {
    editor,
    store,
    access,
    adapter,
    direct,
    deps,
    original,
    insertBreak,
    context,
    bytes: async () => (await buildDocBytes(context()))!,
    lock: () => {
      locked = true
      editor.setEditable(false)
    },
    rebase: async () => {
      doc = { ...doc, parsed: await parseDocx(original) }
    },
  }
}
const prepare = async (adapter: WorkspaceDocumentAdapter, mode: 'edit' | 'reference' = 'edit') =>
  adapter.prepare({ runId: `layout-run-${++sequence}`, mode, state: await adapter.capture() })

afterEach(() => {
  for (const adapter of adapters.splice(0)) adapter.dispose()
  for (const editor of editors.splice(0)) editor.destroy()
  endOdtSession()
  vi.useRealTimers()
})

describe('workspace layout stores', () => {
  it('matches the original long-form writer save bytes and restores the exact pre-write bytes', async () => {
    vi.useFakeTimers({ toFake: ['Date'] })
    vi.setSystemTime(new Date('2026-11-05T00:00:00Z'))
    const direct = await fixture()
    const adapted = await fixture(direct.original)
    const before = await adapted.bytes()
    const html =
      '<h1>Fixture report</h1><p>Factual paragraph.</p><table><tr><td><p>Cell one</p></td><td><p>Cell two</p></td></tr></table>'
    const tool = call('write_document', { plan: 'A report', replaceDocument: true })
    const expected = await executeTool(
      direct.editor,
      tool,
      NUM_IDS,
      undefined,
      undefined,
      undefined,
      undefined,
      undefined,
      {
        async write() {
          return { ok: true, html }
        },
      },
    )
    const run = await prepare(adapted.adapter)
    const executing = run.execute(tool)
    await adapted.adapter.updateWriter!(tool.id, { kind: 'ready' })
    await adapted.adapter.updateWriter!(tool.id, { kind: 'result', ok: true, text: html })
    const actual = await executing
    expect(actual.result).toEqual(expected)
    expect(await adapted.bytes()).toEqual(await direct.bytes())
    run.release()
    await adapted.adapter.rollback!(run.rollbackId!, actual.state)
    expect(await adapted.bytes()).toEqual(before)
  })
  it('exports the current unsaved body and layout through the existing serializer without changing the save baseline', async () => {
    const f = await fixture()
    const run = await prepare(f.adapter)
    await run.execute(call('set_header_footer', { kind: 'header', text: 'Template header' }))
    await run.execute(call('set_page_setup', { orientation: 'landscape' }))
    run.release()
    const state = await f.adapter.capture()
    const before = f.context().doc
    const snapshot = await f.adapter.template!(state, 'as-is')
    expect(snapshot).toMatchObject({ fileName: 'fixture.docx', title: 'fixture', redacted: false })
    expect(snapshot!.bytes).toEqual(await f.bytes())
    expect(f.context().doc).toBe(before)
    expect((await f.adapter.capture()).revision).toBe(state.revision)
  })
  it('matches the existing executor and save bytes for header variants, page setup, styles and watermark in a batch', async () => {
    vi.useFakeTimers({ toFake: ['Date'] })
    vi.setSystemTime(new Date('2026-11-05T00:00:00Z'))
    const direct = await fixture()
    const adapted = await fixture(direct.original)
    direct.direct.buildContext?.()
    const beforeBytes = await adapted.bytes()
    const session = await prepare(adapted.adapter)
    for (const tool of [
      call('set_header_footer', { kind: 'header', text: 'HEADER_ONE' }),
      call('set_header_footer', { kind: 'header', text: 'HEADER_TWO' }),
      call('set_header_footer', { kind: 'footer', text: 'FOOTER' }),
      call('set_header_footer', { kind: 'header', view: 'first', text: 'FIRST_HEADER' }),
      call('set_header_footer', { kind: 'footer', view: 'even', text: 'EVEN_FOOTER' }),
      call('set_page_setup', { paper: 'A4', orientation: 'landscape', margins: { left: '2cm' } }),
      call('set_page_setup', { margins: { right: '3cm' } }),
      call('define_style', {
        styleId: 'WorkspaceQuote',
        name: 'Workspace quote',
        basedOn: 'Normal',
        run: { fontSize: 16 },
      }),
      call('define_style', { styleId: 'WorkspaceQuote', run: { bold: true } }),
      call('apply_ops', {
        ops: [{ op: 'applyStyle', target: { blockIndexes: [0] }, styleId: 'WorkspaceQuote' }],
      }),
      call('set_watermark', { text: 'WORKSPACE_DRAFT' }),
      call('get_document_context'),
      call('list_styles'),
    ]) {
      const revision = adapted.store.revision
      const expected = await direct.direct.executeTool(tool)
      const actual = await session.execute(tool)
      expect(actual.result.isError, tool.name + ': ' + actual.result.output).not.toBe(true)
      expect(actual.result, tool.name).toEqual({
        ...expected,
        mutated: expected.mutated || revision !== adapted.store.revision,
      })
      expect(adapted.store.getSnapshot(), tool.name).toEqual(direct.store.getSnapshot())
      expect(adapted.editor.getJSON(), tool.name).toEqual(direct.editor.getJSON())
    }
    expect(adapted.access.hf.read()).toMatchObject({
      header: 'HEADER_TWO',
      headerFirst: 'FIRST_HEADER',
      footerEven: 'EVEN_FOOTER',
    })
    const bytes = await adapted.bytes()
    expect(bytes).toEqual(await direct.bytes())
    const saved = await parseDocx(bytes)
    expect(saved.styles.get('WorkspaceQuote')).toBeDefined()
    expect(readSectionSettings(saved)?.marginLeft).toBe(1134)
    expect(saved.headerText).toBe('HEADER_TWO')
    expect(adapted.original).toEqual(direct.original)
    session.release()
    await adapted.adapter.rollback!(session.rollbackId!, session.state)
    expect(await adapted.bytes()).toEqual(beforeBytes)
  })

  it('restores store-only edits without replacing the body, then refreshes the original selection for another request', async () => {
    const { editor, store, adapter, access, bytes } = await fixture()
    const beforeNode = editor.state.doc
    const before = store.capture()
    const pinned = await adapter.capture()
    const session = await prepare(adapter)
    for (const tool of [
      call('set_header_footer', { kind: 'header', text: 'AI header' }),
      call('set_header_footer', { kind: 'header', view: 'first', text: 'AI first header' }),
      call('set_header_footer', { kind: 'footer', view: 'even', text: 'AI even footer' }),
      call('set_page_setup', { orientation: 'landscape' }),
      call('define_style', { styleId: 'WorkspaceStyle', name: 'AI style', basedOn: 'Normal' }),
      call('set_watermark', { text: 'AI watermark' }),
    ])
      expect((await session.execute(tool)).result.mutated).toBe(true)
    expect(editor.state.doc).toBe(beforeNode)
    const changedRevision = session.state.revision
    session.release()
    await adapter.rollback!(session.rollbackId!, session.state)
    expect(editor.state.doc).toBe(beforeNode)
    expect(store.getSnapshot()).toMatchObject({
      section: before.section,
      sections: before.sections,
      sectionHfEdits: {},
      styleUpserts: {},
      watermark: null,
      titlePg: false,
      evenOddHf: false,
    })
    const refreshed = await adapter.refresh(pinned)
    expect(refreshed.revision).not.toBe(changedRevision)
    expect(refreshed.selectionId).toBe(pinned.selectionId)
    const next = await adapter.prepare({
      runId: 'after-layout-undo',
      mode: 'edit',
      state: refreshed,
    })
    expect(next.context).not.toContain('AI header')
    expect(access.hf.read().header).toBe('')
    const saved = await parseDocx(await bytes())
    expect(saved.styles.has('WorkspaceStyle')).toBe(false)
    expect(readSectionSettings(saved)).toEqual(before.section)
  })

  it.each(['header', 'sections', 'styleUpserts', 'watermark', 'hfLinks', 'defaultFonts'] as const)(
    'refuses queued work and undo after manual %s changes',
    async (key) => {
      const { adapter, store } = await fixture()
      const pinned = await adapter.capture()
      const session = await prepare(adapter)
      await session.execute(call('set_header_footer', { kind: 'header', text: 'AI header' }))
      session.release()
      if (key === 'header') store.setHeader({ text: 'Manual header' })
      if (key === 'sections')
        store.setSections((sections) => sections.map((sec) => ({ ...sec, titlePg: true })))
      if (key === 'styleUpserts')
        store.setStyleUpserts({ Manual: { styleId: 'Manual', name: 'Manual' } })
      if (key === 'watermark') store.setWatermark('Manual watermark')
      if (key === 'hfLinks') store.setHfLinks({ '0:header': true })
      if (key === 'defaultFonts') store.setDefaultFonts({ font: 'serif' })
      const manual = store.getSnapshot()
      await expect(
        adapter.prepare({ runId: 'queued-layout', mode: 'edit', state: pinned }),
      ).rejects.toMatchObject({ kind: 'stale-document' })
      await expect(adapter.rollback!(session.rollbackId!, session.state)).rejects.toMatchObject({
        kind: 'stale-document',
      })
      expect(store.getSnapshot()).toBe(manual)
    },
  )

  it('retains the existing ODT watermark and text-column restrictions', async () => {
    const Native = globalThis.TextEncoder
    class RealmEncoder extends Native {
      encode(input?: string): Uint8Array<ArrayBuffer> {
        return new Uint8Array(super.encode(input))
      }
    }
    globalThis.TextEncoder = RealmEncoder
    let converted: Awaited<ReturnType<typeof odtToDocx>>
    try {
      converted = await odtToDocx(await buildOdt({ body: hPara('ODT layout fixture') }))
    } finally {
      globalThis.TextEncoder = Native
    }
    beginOdtSession({ bridge: converted.bridge }, await parseDocx(converted.docxBytes))
    const item = await fixture(converted.docxBytes)
    const session = await prepare(item.adapter)
    expect(session.tools.map((tool) => tool.name)).not.toContain('set_watermark')
    await expect(
      session.execute(call('set_watermark', { text: 'Forbidden' })),
    ).rejects.toMatchObject({ kind: 'unknown-tool' })
    const before = item.store.getSnapshot()
    const result = await session.execute(call('set_page_setup', { columns: { count: 2 } }))
    expect(result.result.isError).toBe(true)
    expect(result.result.output).toContain('Text columns')
    expect(item.store.getSnapshot()).toBe(before)
  })

  it('requires an open save baseline before enabling an observed layout store', async () => {
    const item = await fixture()
    const unavailable = createDocsWorkspaceAdapter({ ...item.deps, getLayoutBaseline: undefined })
    adapters.push(unavailable)
    await expect(unavailable.capture()).rejects.toMatchObject({ kind: 'not-ready' })
  })

  it('invalidates rollback when a save changes the parsed baseline even if the body stays identical', async () => {
    const { adapter, rebase } = await fixture()
    const session = await prepare(adapter)
    await session.execute(call('define_style', { styleId: 'WorkspaceStyle', basedOn: 'Normal' }))
    session.release()
    await rebase()
    await expect(adapter.rollback!(session.rollbackId!, session.state)).rejects.toMatchObject({
      kind: 'stale-document',
    })
  })

  it('observes content changes synchronously, ignores dirty-flag clearing and unsubscribes on disposal', async () => {
    const { adapter, store } = await fixture()
    const notify = vi.fn()
    adapter.subscribeChanges(notify)
    const before = await adapter.capture()
    store.setHeaderDirty(true)
    store.setHeaderDirty(false)
    store.setSectionsDirty([0])
    expect(notify).not.toHaveBeenCalled()
    expect((await adapter.capture()).revision).toBe(before.revision)
    store.setHeader({ text: 'Manual' })
    expect(notify).toHaveBeenCalledOnce()
    expect((await adapter.capture()).revision).not.toBe(before.revision)
    adapter.dispose()
    store.setHeader(null)
    expect(notify).toHaveBeenCalledOnce()
  })

  it.each(['reference', 'read-only'] as const)(
    'withholds layout writes in %s mode and keeps reference reads limited to the body',
    async (mode) => {
      const { adapter, store, lock } = await fixture()
      store.setHeader({ text: 'PRIVATE_HEADER' })
      if (mode === 'read-only') lock()
      const session = await prepare(adapter, mode === 'reference' ? 'reference' : 'edit')
      expect(session.tools.map((t) => t.name)).toContain('list_styles')
      for (const name of [
        'set_header_footer',
        'set_page_setup',
        'define_style',
        'set_watermark',
        'insert_section_break',
      ]) {
        expect(session.tools.map((t) => t.name)).not.toContain(name)
        await expect(session.execute(call(name))).rejects.toMatchObject({ kind: 'unknown-tool' })
      }
      session.release()
      expect(
        (await adapter.read(await adapter.capture(), { scope: 'document' })).text,
      ).not.toContain('PRIVATE_HEADER')
    },
  )

  it('rechecks layout changes after the asynchronous first-write hook and rejects late read-only writes', async () => {
    const item = await fixture()
    const guarded = createDocsWorkspaceAdapter({
      ...item.deps,
      beforeFirstWrite: async () => {
        item.store.setHeader({ text: 'Manual change during preparation' })
      },
    })
    adapters.push(guarded)
    const session = await prepare(guarded)
    await expect(
      session.execute(call('define_style', { styleId: 'TooLate' })),
    ).rejects.toMatchObject({ kind: 'stale-document' })
    expect(item.store.getSnapshot().styleUpserts).toEqual({})
    session.release()
    const another = await prepare(item.adapter)
    item.lock()
    await expect(
      another.execute(call('set_page_setup', { orientation: 'landscape' })),
    ).rejects.toMatchObject({ kind: 'read-only' })
  })
})
