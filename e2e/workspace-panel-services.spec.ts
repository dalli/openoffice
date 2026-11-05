import { test, expect } from '@playwright/test'
import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import JSZip from 'jszip'
import { build } from 'vite'
import { launchShell, waitForPageWithUrl, closeAndSaveVideo } from './helpers'
import type { WorkspacePanelServicesFixture } from './fixtures/workspace-panel-services'

declare global {
  interface Window {
    workspaceServicesModule: typeof import('./fixtures/workspace-panel-services')
    workspaceServices: WorkspacePanelServicesFixture
  }
}

test('production shell IPC connects real panel controls, unsaved references, editing, history, stop and rollback', async () => {
  test.setTimeout(120_000)
  const scratch = await mkdtemp(join(tmpdir(), 'ieum-workspace-services-'))
  await build({
    configFile: false,
    logLevel: 'error',
    esbuild: { jsx: 'automatic' },
    define: { 'process.env.NODE_ENV': '"production"' },
    build: {
      outDir: scratch,
      emptyOutDir: false,
      lib: {
        entry: resolve('e2e/fixtures/workspace-panel-services.tsx'),
        formats: ['iife'],
        name: 'workspaceServicesModule',
        cssFileName: 'services',
        fileName: () => 'services.js',
      },
    },
  })
  const source = join(scratch, 'source.md')
  const target = join(scratch, 'target.docx')
  await writeFile(source, 'Stored source\n')
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
  )
  zip.file(
    '_rels/.rels',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  )
  zip.file(
    'word/document.xml',
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Original target</w:t></w:r></w:p></w:body></w:document>',
  )
  const targetBytes = await zip.generateAsync({ type: 'nodebuffer' })
  await writeFile(target, targetBytes)
  const launched = await launchShell({
    onboardingSeen: true,
    openFile: source,
    openFiles: [target],
    recordVideo: false,
    videoDir: 'workspace-panel-services',
  })
  try {
    await launched.app.evaluate(({ ipcMain }) => {
      const scope = globalThis as { workspaceModelCalls?: number }
      scope.workspaceModelCalls = 0
      for (const channel of ['ai:stream', 'ai:chat']) {
        ipcMain.removeHandler(channel)
        ipcMain.handle(channel, () => {
          scope.workspaceModelCalls!++
          throw new Error('Real models disabled in workspace services test')
        })
      }
    })
    const shellPage = await waitForPageWithUrl(launched.app, '/shell/out/renderer/index.html')
    const markdown = await waitForPageWithUrl(launched.app, '://markdown/')
    const docs = await waitForPageWithUrl(launched.app, '://docs/')
    await expect(markdown.locator('.ProseMirror').first()).toBeVisible()
    await expect(docs.locator('.ProseMirror').first()).toBeVisible()
    await markdown.locator('.ProseMirror').first().fill('UNSAVED service reference')
    await shellPage.addStyleTag({
      content: await readFile(join(scratch, 'services.css'), 'utf8'),
    })
    // Install test code through the debugger; the production page keeps script-src 'self'.
    await launched.app.evaluate(
      async ({ webContents }, { url, source }) => {
        const contents = webContents.getAllWebContents().find((wc) => wc.getURL() === url)
        if (!contents) throw new Error('Missing shell page')
        await contents.executeJavaScript(source)
      },
      {
        url: shellPage.url(),
        source: `${await readFile(join(scratch, 'services.js'), 'utf8')}\nwindow.workspaceServicesModule = workspaceServicesModule; void 0;`,
      },
    )
    await shellPage.evaluate(async () => {
      window.workspaceServices =
        await window.workspaceServicesModule.installWorkspacePanelServices()
    })
    const panel = shellPage.locator('.workspace-ai-surface')
    await expect(panel).toBeVisible()
    await panel.getByRole('button', { name: 'New conversation', exact: true }).click()
    await expect
      .poll(() =>
        shellPage.evaluate(() => {
          const f = window.workspaceServices
          return { selected: f.selected() !== null, error: f.controller.hasOpenError() }
        }),
      )
      .toEqual({ selected: true, error: false })
    const ids = await shellPage.evaluate(() => {
      const f = window.workspaceServices
      const documents = f.controller.documentList()
      return {
        p: f.selected(),
        target: documents.find((d) => d.kind === 'docs')!.id,
        source: documents.find((d) => d.kind === 'markdown')!.id,
      }
    })
    await panel.getByRole('combobox', { name: 'Edit target', exact: true }).selectOption(ids.target)
    await panel
      .getByRole('combobox', { name: 'Reference document', exact: true })
      .selectOption(ids.source)
    await panel.getByRole('button', { name: 'Add reference', exact: true }).click()
    await expect
      .poll(() =>
        shellPage.evaluate(
          () =>
            window.workspaceServices.controller.selection(window.workspaceServices.selected())
              .references.length,
        ),
      )
      .toBe(1)
    await panel
      .getByRole('textbox', { name: 'Ask about the selected document…' })
      .fill('Apply reference to target')
    await panel.getByRole('button', { name: 'Send', exact: true }).click()
    await expect
      .poll(() => shellPage.evaluate(() => window.workspaceServices.requests.length))
      .toBe(1)
    expect(
      await shellPage.evaluate(() => JSON.stringify(window.workspaceServices.requests[0])),
    ).toContain('UNSAVED service reference')
    await panel.getByRole('button', { name: 'New conversation', exact: true }).click()
    await shellPage.evaluate(() => {
      const cb = window.workspaceServices.callbacks[0]!
      cb.onToolCall({
        id: 'apply',
        name: 'replace_blocks',
        input: { startBlockIndex: 0, endBlockIndex: 0, html: '<p>UNSAVED service reference</p>' },
      })
      cb.onDone()
    })
    await expect
      .poll(() => shellPage.evaluate(() => window.workspaceServices.requests.length))
      .toBe(2)
    await shellPage.evaluate(() => {
      const cb = window.workspaceServices.callbacks[1]!
      cb.onDelta('Finished original conversation')
      cb.onDone()
    })
    await expect
      .poll(() => shellPage.evaluate(() => window.workspaceServices.host.busy()))
      .toBe(false)
    await expect(docs.locator('.ProseMirror').first()).toContainText('UNSAVED service reference')
    expect(
      await shellPage.evaluate(
        () =>
          window.workspaceServices.host.conversations.state(window.workspaceServices.selected())
            .messages,
      ),
    ).toEqual([])
    await panel.getByRole('combobox', { name: 'Conversation', exact: true }).selectOption(ids.p)
    await expect(panel).toContainText('Finished original conversation')
    await panel.getByRole('button', { name: 'Undo this run', exact: true }).click()
    await expect(docs.locator('.ProseMirror').first()).toContainText('Original target')
    await expect(panel).toContainText('Reverted')
    await panel.getByRole('combobox', { name: 'Edit target', exact: true }).selectOption(ids.target)
    await panel
      .getByRole('textbox', { name: 'Ask about the selected document…' })
      .fill('Wait for source change')
    await panel.getByRole('button', { name: 'Send', exact: true }).click()
    await expect
      .poll(() => shellPage.evaluate(() => window.workspaceServices.requests.length))
      .toBe(3)
    await markdown.locator('.ProseMirror').first().fill('REFERENCE CHANGED')
    await expect
      .poll(() => shellPage.evaluate(() => window.workspaceServices.cancellations()))
      .toBe(1)
    await expect
      .poll(() => shellPage.evaluate(() => window.workspaceServices.host.busy()))
      .toBe(false)
    await shellPage.evaluate(() => {
      const cb = window.workspaceServices.callbacks[2]!
      cb.onToolCall({
        id: 'late',
        name: 'replace_blocks',
        input: { startBlockIndex: 0, endBlockIndex: 0, html: '<p>must never apply</p>' },
      })
      cb.onDone()
    })
    await expect(docs.locator('.ProseMirror').first()).toContainText('Original target')
    expect(await readFile(source, 'utf8')).toBe('Stored source\n')
    expect(await readFile(target)).toEqual(targetBytes)
    expect(
      await launched.app.evaluate(
        () => (globalThis as { workspaceModelCalls?: number }).workspaceModelCalls,
      ),
    ).toBe(0)
    const reopened = await shellPage.evaluate(async (id) => {
      const f = window.workspaceServices
      return f.client.load(f.host.conversations.state(id).origin)
    }, ids.p)
    expect(reopened.messages.some((m) => m.text === 'Finished original conversation')).toBe(true)
    // Exercise the original document host's pending buffer and first-save IPC result.
    const originalRef = await docs.evaluate(async () => {
      const ref = await window.projectApi.resolveChat({
        filePath: null,
        tempChatId: 'unsaved-workspace-e2e',
      })
      await window.projectApi.appendChat({
        ...ref,
        role: 'user',
        text: 'original pending document question',
      })
      return ref
    })
    await shellPage.evaluate(async (ref) => {
      const f = window.workspaceServices
      await f.refresh()
      const source = f.controller
        .sourceList()
        .find((s) => s.origin.kind === 'legacy-document' && s.origin.chatId === ref.chatId)!
      await f.controller.open(source)
      const id = f.selected()
      f.host.conversations.edit(id, {
        composer: { ...f.host.conversations.state(id).composer, text: 'draft survives first save' },
      })
      f.host.conversations.edit(id, { scrollTop: 47 })
    }, originalRef)
    const originalId = await shellPage.evaluate(() => window.workspaceServices.selected())
    const savedRef = await docs.evaluate(
      async ({ ref, path }) =>
        window.projectApi.rebindChat({
          projectId: ref.projectId,
          tempChatId: ref.chatId,
          newFilePath: path,
        }),
      { ref: originalRef, path: join(scratch, 'history-reference.docx') },
    )
    await expect
      .poll(() =>
        shellPage.evaluate(
          (id) => window.workspaceServices.host.conversations.state(id).origin,
          originalId,
        ),
      )
      .toEqual({ kind: 'legacy-document', ...savedRef })
    expect(
      await shellPage.evaluate((id) => {
        const f = window.workspaceServices
        const state = f.host.conversations.state(id)
        return { selected: f.selected(), draft: state.composer.text, scroll: state.scrollTop }
      }, originalId),
    ).toEqual({ selected: originalId, draft: 'draft survives first save', scroll: 47 })
    await shellPage.evaluate(async (ref) => {
      await window.workspaceServices.client.append(
        { kind: 'legacy-document', ...ref },
        { role: 'assistant', text: 'continued after first save' },
      )
    }, originalRef)
    expect(
      (await docs.evaluate((ref) => window.projectApi.loadChat(ref), savedRef)).at(-1)?.text,
    ).toBe('continued after first save')

    // The real home component owns this transcript and receives the workspace append.
    await shellPage.evaluate(() => window.aiOfficeTabs.activate('home'))
    await shellPage.getByRole('button', { name: 'Ask AI', exact: true }).click()
    await expect
      .poll(() =>
        shellPage.evaluate(async () =>
          (await window.workspaceServices.client.catalog()).sources.some(
            (s) => s.origin.kind === 'window-memory',
          ),
        ),
      )
      .toBe(true)
    await shellPage.evaluate(async () => {
      const f = window.workspaceServices
      await f.refresh()
      const source = f.controller.sourceList().find((s) => s.origin.kind === 'window-memory')!
      await f.controller.open(source)
      await f.client.append(source.origin, {
        role: 'assistant',
        text: 'original home receives workspace reply',
      })
    })
    await expect(shellPage.locator('.kb-ask-msg')).toContainText(
      'original home receives workspace reply',
    )
    expect(
      await launched.app.evaluate(
        () => (globalThis as { workspaceModelCalls?: number }).workspaceModelCalls,
      ),
    ).toBe(0)
    await shellPage.evaluate(() => window.workspaceServices.dispose())
    await expect(panel).toHaveCount(0)
  } finally {
    await closeAndSaveVideo(launched, 'workspace-panel-services')
  }
})
