import { test, expect, type ElectronApplication } from '@playwright/test'
import { copyFile, mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { join } from 'node:path'
import { tmpdir } from 'node:os'
import { randomUUID } from 'node:crypto'
import JSZip from 'jszip'
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
