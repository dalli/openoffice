import { test, expect, type ElectronApplication } from '@playwright/test'
import { copyFile, mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { randomUUID } from 'node:crypto'
import JSZip from 'jszip'
import { build } from 'vite'
import { resolve } from 'node:path'
import type {
  WorkspaceExecutionController,
  WorkspaceRunHandle,
} from '../apps/shell/src/main/workspace-ai/execution-controller'
import type { WorkspaceEditorConnection } from '../apps/shell/src/main/workspace-ai/editor-connection'
import type {
  WorkspaceEditorOperation,
  WorkspaceEditorReply,
  WorkspaceEditorValues,
} from '../packages/agent-core/src/workspace-editor-protocol'
import { closeAndSaveVideo, launchShell, waitForPageWithUrl } from './helpers'

/** Exercise the production preload/renderer endpoint from main, without a model or public MCP. */
async function request<O extends WorkspaceEditorOperation>(
  app: ElectronApplication,
  kind: string,
  operation: O,
): Promise<WorkspaceEditorValues[O['kind']]> {
  const reply = await app.evaluate(
    ({ ipcMain, webContents }, message) =>
      new Promise<WorkspaceEditorReply>((resolve, reject) => {
        const contents = webContents
          .getAllWebContents()
          .find((candidate) => candidate.getURL().includes(`://${message.kind}/`))
        if (!contents) {
          reject(new Error('missing editor'))
          return
        }
        const timer = setTimeout(() => {
          ipcMain.removeListener('workspace-ai:editor-reply', listener)
          reject(new Error('editor request timeout'))
        }, 15_000)
        const listener = (event: Electron.IpcMainEvent, reply: WorkspaceEditorReply) => {
          if (
            event.sender !== contents ||
            event.senderFrame !== contents.mainFrame ||
            reply.id !== message.id
          )
            return
          clearTimeout(timer)
          ipcMain.removeListener('workspace-ai:editor-reply', listener)
          resolve(reply)
        }
        ipcMain.on('workspace-ai:editor-reply', listener)
        contents.send('workspace-ai:editor-request', {
          id: message.id,
          operation: message.operation,
        })
      }),
    { id: randomUUID(), kind, operation },
  )
  if (!reply.ok) throw new Error(`workspace adapter: ${reply.error}`)
  return reply.value as WorkspaceEditorValues[O['kind']]
}

async function minimalDocx(): Promise<Buffer> {
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
  return zip.generateAsync({ type: 'nodebuffer' })
}

async function blockModels(app: ElectronApplication): Promise<void> {
  await app.evaluate(({ ipcMain }) => {
    const scope = globalThis as { __workspaceModelCalls?: number }
    scope.__workspaceModelCalls = 0
    for (const channel of ['ai:stream', 'ai:chat']) {
      ipcMain.removeHandler(channel)
      ipcMain.handle(channel, () => {
        scope.__workspaceModelCalls = (scope.__workspaceModelCalls ?? 0) + 1
        throw new Error('Models are disabled in workspace adapter tests')
      })
    }
  })
}

test('reads unsaved Markdown into DOCX, preserving reference, disk bytes, undo and stale-write rejection', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'ieum-workspace-editors-'))
  const source = join(dir, 'source.md')
  const target = join(dir, 'target.docx')
  await writeFile(source, 'Stored reference\n')
  const targetBytes = await minimalDocx()
  await writeFile(target, targetBytes)
  const launched = await launchShell({
    onboardingSeen: true,
    openFile: source,
    openFiles: [target],
    recordVideo: false,
    videoDir: 'workspace-adapters',
  })
  try {
    await blockModels(launched.app)
    const markdown = await waitForPageWithUrl(launched.app, '://markdown/')
    const docs = await waitForPageWithUrl(launched.app, '://docs/')
    await expect(markdown.locator('.ProseMirror').first()).toBeVisible()
    await expect(docs.locator('.ProseMirror').first()).toBeVisible()
    await markdown.locator('.ProseMirror').first().fill('UNSAVED reference marker')
    const sourceState = await request(launched.app, 'markdown', { kind: 'capture' })
    const material = await request(launched.app, 'markdown', {
      kind: 'read',
      state: sourceState,
      read: { scope: 'document' },
    })
    expect(material.text).toContain('UNSAVED reference marker')
    expect(material.masking).toBe('applied')
    const reference = await request(launched.app, 'markdown', {
      kind: 'prepare',
      request: { runId: 'reference-P', mode: 'reference', state: sourceState },
    })
    expect(reference.tools.some((tool) => tool.name === 'apply_ops')).toBe(false)

    const targetState = await request(launched.app, 'docs', { kind: 'capture' })
    const session = await request(launched.app, 'docs', {
      kind: 'prepare',
      request: { runId: 'target-P', mode: 'edit', state: targetState },
    })
    const applied = await request(launched.app, 'docs', {
      kind: 'execute',
      sessionId: session.id,
      runId: session.runId,
      call: {
        id: 'apply-1',
        name: 'replace_blocks',
        input: { startBlockIndex: 0, endBlockIndex: 0, html: `<p>${material.text.trim()}</p>` },
      },
    })
    expect(applied.result.mutated).toBe(true)
    await expect(docs.locator('.ProseMirror').first()).toContainText('UNSAVED reference marker')
    expect(
      (
        await request(launched.app, 'markdown', {
          kind: 'read',
          state: sourceState,
          read: { scope: 'document' },
        })
      ).text,
    ).toBe(material.text)
    expect(await readFile(source, 'utf8')).toBe('Stored reference\n')
    expect(await readFile(target)).toEqual(targetBytes)

    await request(launched.app, 'docs', {
      kind: 'release',
      sessionId: session.id,
      runId: session.runId,
    })
    await docs
      .locator('.ProseMirror')
      .first()
      .press(process.platform === 'darwin' ? 'Meta+z' : 'Control+z')
    await expect(docs.locator('.ProseMirror').first()).toContainText('Original target')
    const current = await request(launched.app, 'docs', { kind: 'capture' })
    const stale = await request(launched.app, 'docs', {
      kind: 'prepare',
      request: { runId: 'stale-P', mode: 'edit', state: current },
    })
    await docs.locator('.ProseMirror').first().fill('Manual change wins')
    await expect(
      request(launched.app, 'docs', {
        kind: 'execute',
        sessionId: stale.id,
        runId: stale.runId,
        call: { id: 'late-1', name: 'insert_content', input: { html: '<p>Late result</p>' } },
      }),
    ).rejects.toThrow('stale-document')
    await expect(docs.locator('.ProseMirror').first()).toHaveText('Manual change wins')
    expect(
      await launched.app.evaluate(
        () => (globalThis as { __workspaceModelCalls?: number }).__workspaceModelCalls,
      ),
    ).toBe(0)
  } finally {
    await closeAndSaveVideo(launched, 'workspace-adapters')
  }
})

test('runs the HWP plugin version guard through the built preload and preserves the disk file', async () => {
  const dir = await mkdtemp(join(tmpdir(), 'ieum-workspace-hwp-'))
  const path = join(dir, 'reference.hwpx')
  await copyFile(join(__dirname, '../apps/hwp/tests/fixtures/memo.hwpx'), path)
  const before = await readFile(path)
  const launched = await launchShell({
    onboardingSeen: true,
    openFile: path,
    recordVideo: false,
    videoDir: 'workspace-hwp-adapter',
  })
  try {
    await blockModels(launched.app)
    const page = await waitForPageWithUrl(launched.app, '://hwp/')
    await page.waitForSelector('.hwp-app[data-phase="ready"]')
    const state = await request(launched.app, 'hwp', { kind: 'capture' })
    const initial = await request(launched.app, 'hwp', {
      kind: 'read',
      state,
      read: { scope: 'document' },
    })
    expect(initial.masking).toBe('unsupported')
    const session = await request(launched.app, 'hwp', {
      kind: 'prepare',
      request: { runId: 'hwp-P', mode: 'edit', state },
    })
    const applied = await request(launched.app, 'hwp', {
      kind: 'execute',
      sessionId: session.id,
      runId: session.runId,
      call: {
        id: 'hwp-1',
        name: 'insert_paragraphs',
        input: { where: 'end', paragraphs: ['UNSAVED HWP marker'] },
      },
    })
    expect(applied.result.mutated).toBe(true)
    const latest = await request(launched.app, 'hwp', {
      kind: 'read',
      state: applied.state,
      read: { scope: 'document' },
    })
    expect(latest.text).toContain('UNSAVED HWP marker')
    await expect(
      request(launched.app, 'hwp', { kind: 'read', state, read: { scope: 'document' } }),
    ).rejects.toThrow('stale-document')
    expect(await readFile(path)).toEqual(before)
    expect(
      await launched.app.evaluate(
        () => (globalThis as { __workspaceModelCalls?: number }).__workspaceModelCalls,
      ),
    ).toBe(0)
  } finally {
    await closeAndSaveVideo(launched, 'workspace-hwp-adapter')
  }
})

interface MainWorkspaceHarness {
  controller: WorkspaceExecutionController
  clients: WorkspaceEditorConnection[]
  run: WorkspaceRunHandle
  target: import('../apps/shell/src/shared/workspace-ai-api').WorkspaceDocumentIdentity
  referenceText: string
  queuedHwp: WorkspaceRunHandle
  signal: AbortSignal
}

test('coordinates the FIFO registry and production editor clients with explicit references and invalidation', async () => {
  test.setTimeout(120_000)
  const dir = await mkdtemp(join(tmpdir(), 'ieum-workspace-controller-'))
  await build({
    configFile: false,
    logLevel: 'error',
    build: {
      outDir: dir,
      emptyOutDir: false,
      target: 'node22',
      lib: {
        entry: resolve('e2e/fixtures/workspace-execution-main.ts'),
        formats: ['cjs'],
        fileName: () => 'execution.cjs',
      },
      rollupOptions: { external: ['node:crypto'] },
    },
  })
  const source = join(dir, 'source.md')
  const target = join(dir, 'target.docx')
  const hwp = join(dir, 'third.hwpx')
  await writeFile(source, 'Stored reference\n')
  const targetBytes = await minimalDocx()
  await writeFile(target, targetBytes)
  await copyFile(join(__dirname, '../apps/hwp/tests/fixtures/memo.hwpx'), hwp)
  const hwpBytes = await readFile(hwp)
  const launched = await launchShell({
    onboardingSeen: true,
    openFile: source,
    openFiles: [target, hwp],
    recordVideo: false,
    videoDir: 'workspace-controller',
  })
  try {
    await blockModels(launched.app)
    const markdown = await waitForPageWithUrl(launched.app, '://markdown/')
    const docs = await waitForPageWithUrl(launched.app, '://docs/')
    const hwpPage = await waitForPageWithUrl(launched.app, '://hwp/')
    await expect(markdown.locator('.ProseMirror').first()).toBeVisible()
    await expect(docs.locator('.ProseMirror').first()).toBeVisible()
    await hwpPage.waitForSelector('.hwp-app[data-phase="ready"]')
    await markdown.locator('.ProseMirror').first().fill('UNSAVED material from A')
    await launched.app.evaluate(
      async ({ BrowserWindow, WebContentsView, ipcMain }, modulePath) => {
        const load = process.getBuiltinModule('module').createRequire(modulePath)
        const { WorkspaceExecutionController, WorkspaceEditorConnection } = load(
          modulePath,
        ) as typeof import('./fixtures/workspace-execution-main')
        const win = BrowserWindow.getAllWindows()[0]
        const views = win.contentView.children.filter(
          (view): view is Electron.WebContentsView => view instanceof WebContentsView,
        )
        const documents = new Map<
          string,
          { kind: 'docs' | 'markdown' | 'hwp'; client: WorkspaceEditorConnection }
        >()
        for (const kind of ['docs', 'markdown', 'hwp'] as const) {
          const view = views.find((item) => item.webContents.getURL().includes(`://${kind}/`))!
          documents.set(kind, {
            kind,
            client: new WorkspaceEditorConnection(view.webContents, ipcMain, () =>
              win.contentView.children.includes(view),
            ),
          })
        }
        const controller = new WorkspaceExecutionController({
          owner: { windowId: win.id, panelContentsId: win.webContents.id, panelGeneration: 1 },
          resolve: (id) => documents.get(id) ?? null,
        })
        const target = await controller.capture('docs')
        const reference = await controller.readReference(
          (await controller.capture('markdown')).id,
          { scope: 'document' },
        )
        const third = await controller.capture('hwp')
        const run = controller.enqueue({
          conversationId: 'conversation-P',
          runId: 'P',
          targetId: target.id,
          referenceIds: [reference.id],
          project: null,
        })
        controller.enqueue({
          conversationId: 'conversation-Q',
          runId: 'Q',
          targetId: target.id,
          referenceIds: [],
          project: null,
        })
        const queuedHwp = controller.enqueue({
          conversationId: 'conversation-R',
          runId: 'R',
          targetId: third.id,
          referenceIds: [reference.id],
          project: null,
        })
        const started = await controller.startNext()
        if (!started) throw new Error('missing started run')
        ;(globalThis as { __workspaceHarness?: MainWorkspaceHarness }).__workspaceHarness = {
          controller,
          clients: [...documents.values()].map((item) => item.client),
          run,
          target: run.target.document,
          referenceText: started.references[0].snapshot.text.trim(),
          queuedHwp,
          signal: started.signal,
        }
      },
      join(dir, 'execution.cjs'),
    )
    await launched.app.evaluate(async () => {
      const harness = (globalThis as unknown as { __workspaceHarness: MainWorkspaceHarness })
        .__workspaceHarness
      await harness.controller.execute(harness.run, harness.target, {
        id: 'copy-to-B',
        name: 'replace_blocks',
        input: { startBlockIndex: 0, endBlockIndex: 0, html: `<p>${harness.referenceText}</p>` },
      })
    })
    await expect(docs.locator('.ProseMirror').first()).toContainText('UNSAVED material from A')
    const after = await launched.app.evaluate(async () => {
      const harness = (globalThis as unknown as { __workspaceHarness: MainWorkspaceHarness })
        .__workspaceHarness
      const invalidated = harness.controller.state('Q')
      await harness.controller.complete(harness.run)
      const next = await harness.controller.startNext()
      if (!next) throw new Error('missing queued HWP run')
      harness.run = next.binding
      harness.target = next.binding.target.document
      harness.signal = next.signal
      return { invalidated, next: next.binding.runId }
    })
    expect(after).toEqual({
      invalidated: { status: 'stopped', reason: 'document-changed' },
      next: 'R',
    })
    await markdown.locator('.ProseMirror').first().fill('A changed after R started')
    await expect
      .poll(() =>
        launched.app.evaluate(() => {
          const harness = (globalThis as unknown as { __workspaceHarness: MainWorkspaceHarness })
            .__workspaceHarness
          return { state: harness.controller.state('R'), aborted: harness.signal.aborted }
        }),
      )
      .toEqual({ state: { status: 'stopped', reason: 'document-changed' }, aborted: true })
    const late = await launched.app.evaluate(async () => {
      const harness = (globalThis as unknown as { __workspaceHarness: MainWorkspaceHarness })
        .__workspaceHarness
      try {
        await harness.controller.execute(harness.run, harness.target, {
          id: 'late-to-HWP',
          name: 'insert_paragraphs',
          input: { where: 'end', paragraphs: ['Must not appear'] },
        })
        return 'unexpected success'
      } catch (error) {
        return (error as { kind: string }).kind
      }
    })
    expect(late).toBe('run-not-active')
    await launched.app.evaluate(async () => {
      const harness = (globalThis as unknown as { __workspaceHarness: MainWorkspaceHarness })
        .__workspaceHarness
      const target = await harness.controller.capture('hwp')
      harness.run = harness.controller.enqueue({
        conversationId: 'conversation-S',
        runId: 'S',
        targetId: target.id,
        referenceIds: [],
        project: null,
      })
      const started = await harness.controller.startNext()
      if (!started) throw new Error('missing manual-change probe')
      harness.signal = started.signal
    })
    const studio = hwpPage.frames().find((frame) => frame.url().includes('/studio/'))!
    await studio.evaluate(async () => {
      const surface = (
        window as unknown as {
          rhwpStudio: {
            plugins: {
              invoke(id: string, method: string, args: unknown[]): Promise<{ ok: boolean }>
            }
          }
        }
      ).rhwpStudio
      // An edit outside the workspace session must still invalidate its native document binding.
      const result = await surface.plugins.invoke('ieum-agent', 'insertParagraphs', [
        '',
        null,
        'end',
        ['Native local edit'],
      ])
      if (!result.ok) throw new Error('local plugin edit refused')
    })
    await expect
      .poll(() =>
        launched.app.evaluate(() => {
          const harness = (globalThis as unknown as { __workspaceHarness: MainWorkspaceHarness })
            .__workspaceHarness
          return { state: harness.controller.state('S'), aborted: harness.signal.aborted }
        }),
      )
      .toEqual({ state: { status: 'stopped', reason: 'document-changed' }, aborted: true })
    const hosted = await launched.app.evaluate(
      async (_electron, modulePath) => {
        const load = process.getBuiltinModule('module').createRequire(modulePath)
        const { WorkspaceRunHost } = load(
          modulePath,
        ) as typeof import('./fixtures/workspace-execution-main')
        const { controller } = (
          globalThis as unknown as { __workspaceHarness: MainWorkspaceHarness }
        ).__workspaceHarness
        const target = await controller.capture('docs')
        const reference = await controller.readReference(
          (await controller.capture('markdown')).id,
          { scope: 'document' },
        )
        const before = await controller.readReference(target.id, { scope: 'document' })
        let modelTurns = 0
        const host = new WorkspaceRunHost({
          execution: controller,
          load: async () => ({ messages: [], foreign: false }),
          append: async () => 'memory',
          admit: () => true,
          gate: () => ({
            kind: 'run',
            choice: { connected: true, provider: 'claude-code', metered: false },
          }),
          errorText: (error) => (error instanceof Error ? error.message : String(error)),
          prepare: async () => ({ project: null, skills: [], dispose() {} }),
          createTransport: () => ({
            stream: (request, callbacks) => {
              const turn = ++modelTurns
              queueMicrotask(() => {
                if (turn === 1) {
                  const currentMaterial = request.messages.some(
                    (message) =>
                      message.role === 'user' && message.text.includes('A changed after R started'),
                  )
                  callbacks.onToolCall({
                    id: 'hosted-edit',
                    name: 'replace_blocks',
                    input: {
                      startBlockIndex: 0,
                      endBlockIndex: 0,
                      html: `<p>HOST current reference: ${currentMaterial ? 'A changed after R started' : 'MISSING'}</p>`,
                    },
                  })
                } else callbacks.onDelta('Completed host edit.')
                callbacks.onDone()
              })
              return { cancel() {} }
            },
          }),
        })
        try {
          const p = await host.conversations.open(
            { kind: 'workspace', conversationId: 'host-P' },
            target.binding,
          )
          const q = await host.conversations.open({ kind: 'workspace', conversationId: 'host-Q' })
          host.conversations.edit(p, {
            composer: {
              ...host.conversations.state(p).composer,
              text: 'Apply the current explicit reference to the target.',
              references: [reference.binding],
            },
          })
          const submitted = await host.submit(p, target.id, [reference.id])
          if (!submitted) throw new Error('host did not accept the explicit request')
          host.conversations.select(q)
          await submitted.done
          const changed = await controller.readReference((await controller.capture('docs')).id, {
            scope: 'document',
          })
          const card = host.rollbackCards(p)[0]
          if (!card)
            throw new Error(
              `missing rollback card: ${JSON.stringify(host.conversations.state(p).messages.at(-1))}; ${host.conversations.state(p).error}`,
            )
          await host.rollback(p, card.id)
          const restored = await controller.readReference((await controller.capture('docs')).id, {
            scope: 'document',
          })
          return {
            modelTurns,
            changed: changed.snapshot.text,
            restored: restored.snapshot.text,
            before: before.snapshot.text,
            otherMessages: host.conversations.state(q).messages.length,
            rollback: host.rollbackCards(p)[0]?.state,
            error: host.conversations.state(p).error,
          }
        } finally {
          host.dispose()
        }
      },
      join(dir, 'execution.cjs'),
    )
    expect(hosted.modelTurns).toBe(2)
    expect(hosted.changed).toContain('HOST current reference: A changed after R started')
    expect(hosted.restored).toBe(hosted.before)
    expect(hosted.otherMessages).toBe(0)
    expect(hosted.rollback).toBe('reverted')
    expect(hosted.error).toBeNull()
    expect(await readFile(source, 'utf8')).toBe('Stored reference\n')
    expect(await readFile(target)).toEqual(targetBytes)
    expect(await readFile(hwp)).toEqual(hwpBytes)
    expect(
      await launched.app.evaluate(
        () => (globalThis as { __workspaceModelCalls?: number }).__workspaceModelCalls,
      ),
    ).toBe(0)
  } finally {
    await launched.app
      .evaluate(() => {
        const harness = (globalThis as { __workspaceHarness?: MainWorkspaceHarness })
          .__workspaceHarness
        harness?.controller.dispose()
        harness?.clients.forEach((client) => client.dispose())
      })
      .catch(() => {})
    await closeAndSaveVideo(launched, 'workspace-controller')
  }
})
