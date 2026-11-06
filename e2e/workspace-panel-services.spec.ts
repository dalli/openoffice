import { test, expect } from '@playwright/test'
import { mkdir, mkdtemp, readFile, readdir, writeFile } from 'node:fs/promises'
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
      const scope = globalThis as { workspaceModelCalls?: number; workspaceMcpCalls?: unknown[] }
      scope.workspaceModelCalls = 0
      scope.workspaceMcpCalls = []
      // Only the hub boundary is fake; the panel uses the production MCP preload bridge.
      ipcMain.removeHandler('mcp-tools:list')
      ipcMain.handle('mcp-tools:list', () =>
        ['alpha', 'beta'].map((name) => ({
          name,
          status: 'ready',
          newToolCount: 0,
          keepResultsInside: true,
          instructions: `Approved ${name} instructions`,
          tools: [
            {
              name: 'lookup',
              exposedName: `mcp__${name}__lookup`,
              description: 'Approved lookup',
              inputSchema: { type: 'object' },
              permission: 'ask',
              pending: false,
            },
          ],
        })),
      )
      ipcMain.removeHandler('mcp-tools:call')
      ipcMain.handle('mcp-tools:call', (_event, request) => {
        scope.workspaceMcpCalls!.push(request)
        return { ok: true, output: 'WORKSPACE_MCP_RESULT', isError: false }
      })
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
    // The previous conversation deliberately retains a stale reference; use a fresh attachment draft.
    await panel.getByRole('button', { name: 'New conversation', exact: true }).click()
    await panel.getByRole('combobox', { name: 'Edit target', exact: true }).selectOption(ids.target)
    // Native picker + production parser/image reader, with a test-only model transport.
    const material = join(scratch, 'attached-material.md')
    const image = join(scratch, 'attached-image.png')
    const materialText =
      'ATTACHED VISIBLE <span class="redact-span" data-redaction="" data-label="private">ATTACHED HIDDEN</span>'
    const imageBytes = Buffer.from(
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+a/BcAAAAASUVORK5CYII=',
      'base64',
    )
    await writeFile(material, materialText)
    await writeFile(image, imageBytes)
    await launched.app.evaluate(
      ({ dialog }, paths) => {
        const state = globalThis as { restoreAttachmentDialog?: () => void }
        const original = dialog.showOpenDialog
        Object.assign(dialog, {
          showOpenDialog: async () => ({ canceled: false, filePaths: paths }),
        })
        state.restoreAttachmentDialog = () => {
          dialog.showOpenDialog = original
        }
      },
      [material, image],
    )
    const beforeAttachments = await shellPage.evaluate(
      () => window.workspaceServices.requests.length,
    )
    await panel
      .getByRole('button', {
        name: 'Attach local files (or drag them onto the panel)',
        exact: true,
      })
      .click()
    await expect(panel.locator('.workspace-ai-attachments')).toContainText('attached-material.md')
    await expect(panel.locator('.workspace-ai-attachments')).toContainText('attached-image.png')
    await launched.app.evaluate(() =>
      (globalThis as { restoreAttachmentDialog?: () => void }).restoreAttachmentDialog?.(),
    )
    expect(await shellPage.evaluate(() => window.workspaceServices.requests.length)).toBe(
      beforeAttachments,
    )
    await panel.getByRole('textbox').fill('Read the attached material and image')
    await panel.getByRole('button', { name: 'Send', exact: true }).click()
    await expect
      .poll(() => shellPage.evaluate(() => window.workspaceServices.requests.length))
      .toBe(beforeAttachments + 1)
    const attachmentRequest = await shellPage.evaluate(() =>
      window.workspaceServices.requests.at(-1)!,
    )
    expect(JSON.stringify(attachmentRequest)).toContain(imageBytes.toString('base64'))
    expect(JSON.stringify(attachmentRequest)).toContain('attached-material.md')
    expect(JSON.stringify(attachmentRequest)).not.toContain('ATTACHED HIDDEN')
    await shellPage.evaluate(() => {
      const cb = window.workspaceServices.callbacks.at(-1)!
      cb.onToolCall({ id: 'read-attachment', name: 'read_attachment', input: { index: 0 } })
      cb.onDone()
    })
    await expect
      .poll(() => shellPage.evaluate(() => window.workspaceServices.requests.length))
      .toBe(beforeAttachments + 2)
    const attachmentResult = JSON.stringify(
      await shellPage.evaluate(() => window.workspaceServices.requests.at(-1)!.messages),
    )
    expect(attachmentResult).toContain('ATTACHED VISIBLE')
    expect(attachmentResult).toContain('{{private}}')
    expect(attachmentResult).not.toContain('ATTACHED HIDDEN')
    await shellPage.evaluate(() => {
      const cb = window.workspaceServices.callbacks.at(-1)!
      cb.onDelta('Finished attachment request')
      cb.onDone()
    })
    await expect
      .poll(() => shellPage.evaluate(() => window.workspaceServices.host.busy()))
      .toBe(false)
    const attachmentHistory = await shellPage.evaluate(() => {
      const f = window.workspaceServices
      return f.client.load(f.host.conversations.state(f.selected()).origin)
    })
    expect(attachmentHistory.messages.at(-2)?.attachments?.map((file) => file.name)).toEqual([
      'attached-material.md',
      'attached-image.png',
    ])
    expect(JSON.stringify(attachmentHistory)).not.toContain('ATTACHED VISIBLE')
    expect(JSON.stringify(attachmentHistory)).not.toContain(imageBytes.toString('base64'))
    expect(await readFile(material, 'utf8')).toBe(materialText)
    expect(await readFile(image)).toEqual(imageBytes)
    // A pathless clipboard bitmap uses the existing pasted-image temp-file service.
    await shellPage.evaluate(
      (bytes) => {
        const transfer = new DataTransfer()
        transfer.items.add(
          new File([new Uint8Array(bytes)], 'screenshot.png', { type: 'image/png' }),
        )
        document.querySelector('.workspace-ai-surface textarea')!.dispatchEvent(
          new ClipboardEvent('paste', {
            clipboardData: transfer,
            bubbles: true,
            cancelable: true,
          }),
        )
      },
      [...imageBytes],
    )
    await expect(panel.locator('.workspace-ai-attachments')).toContainText(/pasted-.+\.png/)
    await panel.getByRole('textbox').fill('Read the pasted screenshot')
    await panel.getByRole('button', { name: 'Send', exact: true }).click()
    await expect
      .poll(() => shellPage.evaluate(() => window.workspaceServices.requests.length))
      .toBe(beforeAttachments + 3)
    expect(
      JSON.stringify(await shellPage.evaluate(() => window.workspaceServices.requests.at(-1)!)),
    ).toContain(imageBytes.toString('base64'))
    await shellPage.evaluate(() => {
      const cb = window.workspaceServices.callbacks.at(-1)!
      cb.onDelta('Finished pasted image request')
      cb.onDone()
    })
    await expect
      .poll(() => shellPage.evaluate(() => window.workspaceServices.host.busy()))
      .toBe(false)

    // MCP server choices and permission cards follow the conversation, through the real preload.
    await panel.getByRole('button', { name: 'New conversation', exact: true }).click()
    const mcpId = await shellPage.evaluate(() => window.workspaceServices.selected())
    await panel.getByRole('combobox', { name: 'Edit target', exact: true }).selectOption(ids.target)
    const beforeMcp = await shellPage.evaluate(() => window.workspaceServices.requests.length)
    await panel.getByRole('textbox').fill('Use the connected lookup server')
    await panel.getByRole('button', { name: 'Send', exact: true }).click()
    await expect
      .poll(() => shellPage.evaluate(() => window.workspaceServices.requests.length))
      .toBe(beforeMcp + 1)
    await shellPage.evaluate(() => {
      const cb = window.workspaceServices.callbacks.at(-1)!
      cb.onToolCall({ id: 'mcp-lookup', name: 'mcp__alpha__lookup', input: {} })
      cb.onDone()
    })
    await expect(panel.locator('.ai-mcp-card')).toContainText('alpha')
    expect(
      await launched.app.evaluate(
        () => (globalThis as { workspaceMcpCalls?: unknown[] }).workspaceMcpCalls!.length,
      ),
    ).toBe(0)
    await panel.getByRole('button', { name: 'New conversation', exact: true }).click()
    await expect(panel.locator('.ai-mcp-card')).toHaveCount(0)
    await panel.getByRole('button', { name: 'Connected servers', exact: true }).click()
    await panel
      .getByRole('checkbox', { name: 'Use alpha in this conversation', exact: true })
      .uncheck()
    await panel.locator('.ai-mcp-pop').press('Escape')
    await panel.getByRole('combobox', { name: 'Conversation', exact: true }).selectOption(mcpId)
    await expect(panel.locator('.ai-mcp-card')).toContainText('alpha')
    await panel
      .getByRole('button', { name: 'Always allow in this conversation', exact: true })
      .click()
    await expect
      .poll(() => shellPage.evaluate(() => window.workspaceServices.requests.length))
      .toBe(beforeMcp + 2)
    expect(
      JSON.stringify(await shellPage.evaluate(() => window.workspaceServices.requests.at(-1))),
    ).toContain('WORKSPACE_MCP_RESULT')
    await shellPage.evaluate(() => {
      const cb = window.workspaceServices.callbacks.at(-1)!
      cb.onToolCall({ id: 'other-server', name: 'mcp__beta__lookup', input: {} })
      cb.onDone()
    })
    await expect
      .poll(() => shellPage.evaluate(() => window.workspaceServices.requests.length))
      .toBe(beforeMcp + 3)
    expect(
      await launched.app.evaluate(
        () => (globalThis as { workspaceMcpCalls?: unknown[] }).workspaceMcpCalls!.length,
      ),
    ).toBe(1)
    expect(
      JSON.stringify(await shellPage.evaluate(() => window.workspaceServices.requests.at(-1))),
    ).toContain('set to keep what it returns inside')
    await shellPage.evaluate(() => {
      const cb = window.workspaceServices.callbacks.at(-1)!
      cb.onDelta('Finished MCP request')
      cb.onDone()
    })
    await expect
      .poll(() => shellPage.evaluate(() => window.workspaceServices.host.busy()))
      .toBe(false)
    await panel.getByRole('button', { name: 'Connected servers', exact: true }).click()
    await expect(
      panel.getByRole('checkbox', { name: 'Use alpha in this conversation', exact: true }),
    ).toBeChecked()
    await panel.locator('.ai-mcp-pop').press('Escape')

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
    // Register a synthetic folder through the existing picker and trust dialogs.
    const projectFolder = join(scratch, 'panel-project')
    await mkdir(projectFolder)
    await writeFile(join(projectFolder, 'AGENTS.md'), 'WORKSPACE PROJECT RULES')
    await writeFile(join(projectFolder, 'material.md'), 'PROJECT READ MATERIAL')
    await launched.app.evaluate(({ dialog }, folder) => {
      const saved = globalThis as {
        workspaceDialogs?: {
          open: typeof dialog.showOpenDialog
          message: typeof dialog.showMessageBox
        }
      }
      saved.workspaceDialogs = { open: dialog.showOpenDialog, message: dialog.showMessageBox }
      dialog.showOpenDialog = (async () => ({
        canceled: false,
        filePaths: [folder],
      })) as typeof dialog.showOpenDialog
      dialog.showMessageBox = (async () => ({
        response: 0,
        checkboxChecked: false,
      })) as typeof dialog.showMessageBox
    }, projectFolder)
    let projectId: string
    try {
      projectId = await shellPage.evaluate(async () => {
        const bridge = (
          window as unknown as {
            projectFolderApi: import('@genoffice/project-store').ProjectFolderApi
          }
        ).projectFolderApi
        const added = await bridge.addFolder()
        if (added.status !== 'added') throw new Error(`Could not register project: ${added.status}`)
        return added.project.id
      })
    } finally {
      await launched.app.evaluate(({ dialog }) => {
        const saved = globalThis as {
          workspaceDialogs?: {
            open: typeof dialog.showOpenDialog
            message: typeof dialog.showMessageBox
          }
        }
        dialog.showOpenDialog = saved.workspaceDialogs!.open
        dialog.showMessageBox = saved.workspaceDialogs!.message
        delete saved.workspaceDialogs
      })
    }
    const beforeProject = await shellPage.evaluate(() => window.workspaceServices.requests.length)
    await shellPage.evaluate(() => {
      const bridge = (
        window as unknown as {
          projectFolderApi: import('@genoffice/project-store').ProjectFolderApi
        }
      ).projectFolderApi
      const state = window as unknown as { projectMemoryEvents: string[] }
      state.projectMemoryEvents = []
      bridge.onProjectChanged((event) => {
        if (event.part === 'memory') state.projectMemoryEvents.push(event.projectId)
      })
    })
    await panel.locator('.ai-project-chip').click()
    await shellPage.getByRole('option', { name: /panel-project/ }).click()
    await expect
      .poll(() =>
        shellPage.evaluate(
          () =>
            window.workspaceServices.host.conversations.state(window.workspaceServices.selected())
              .origin.kind,
        ),
      )
      .toBe('project')
    expect(await shellPage.evaluate(() => window.workspaceServices.requests.length)).toBe(
      beforeProject,
    )
    expect(await readdir(projectFolder)).toEqual(['AGENTS.md', 'material.md'])
    await panel.getByRole('combobox', { name: 'Edit target', exact: true }).selectOption(ids.target)
    await panel.getByRole('textbox').fill('Read material.md in the chosen project')
    await panel.getByRole('button', { name: 'Send', exact: true }).click()
    await expect
      .poll(() => shellPage.evaluate(() => window.workspaceServices.requests.length))
      .toBe(beforeProject + 1)
    const projectRequest = await shellPage.evaluate(() => window.workspaceServices.requests.at(-1)!)
    expect(projectRequest.system).toContain('WORKSPACE PROJECT RULES')
    expect(JSON.stringify(projectRequest.messages)).not.toContain('WORKSPACE PROJECT RULES')
    await shellPage.evaluate(() => {
      const cb = window.workspaceServices.callbacks.at(-1)!
      cb.onToolCall({ id: 'read-project', name: 'project_read', input: { path: 'material.md' } })
      cb.onToolCall({
        id: 'remember-project',
        name: 'project_remember',
        input: { op: 'add', text: 'Use concise project answers (2026-11-05)' },
      })
      cb.onDone()
    })
    await expect
      .poll(() => shellPage.evaluate(() => window.workspaceServices.requests.length))
      .toBe(beforeProject + 2)
    expect(
      JSON.stringify(
        await shellPage.evaluate(() => window.workspaceServices.requests.at(-1)!.messages),
      ),
    ).toContain('PROJECT READ MATERIAL')
    await shellPage.evaluate(() => {
      const cb = window.workspaceServices.callbacks.at(-1)!
      cb.onDelta('Finished project request')
      cb.onDone()
    })
    await expect
      .poll(() => shellPage.evaluate(() => window.workspaceServices.host.busy()))
      .toBe(false)
    const projectHistory = await shellPage.evaluate(async (projectId) => {
      const f = window.workspaceServices
      const origin = f.host.conversations.state(f.selected()).origin
      if (origin.kind !== 'project' || origin.projectId !== projectId)
        throw new Error('Wrong project')
      return f.client.load(origin)
    }, projectId)
    expect(projectHistory.messages.at(-1)?.text).toBe('Finished project request')
    expect(JSON.stringify(projectHistory)).not.toContain('PROJECT READ MATERIAL')
    expect(await readFile(join(projectFolder, '.ieumoffice', 'memory.md'), 'utf8')).toContain(
      'Use concise project answers',
    )
    expect(
      await shellPage.evaluate(
        () => (window as unknown as { projectMemoryEvents: string[] }).projectMemoryEvents,
      ),
    ).toContain(projectId)
    expect(await readFile(join(projectFolder, 'material.md'), 'utf8')).toBe('PROJECT READ MATERIAL')

    // Read an app-bundled skill through the existing production preload and main guards.
    await panel.getByRole('button', { name: 'New conversation', exact: true }).click()
    await panel.getByRole('combobox', { name: 'Edit target', exact: true }).selectOption(ids.target)
    const beforeSkills = await shellPage.evaluate(() => window.workspaceServices.requests.length)
    const composer = panel.locator('.ai-input-box textarea')
    await composer.fill('/meeting-minutes')
    await composer.press('ArrowRight')
    await expect(panel.getByRole('option').filter({ hasText: 'meeting-minutes' })).toBeVisible()
    await composer.press('Enter')
    await expect(panel.locator('.ai-skill-chips')).toContainText('meeting-minutes')
    await composer.fill('Write the minutes using the selected skill')
    await panel.getByRole('button', { name: 'Send', exact: true }).click()
    await expect
      .poll(() => shellPage.evaluate(() => window.workspaceServices.requests.length))
      .toBe(beforeSkills + 1)
    expect(
      JSON.stringify(
        await shellPage.evaluate(() => window.workspaceServices.requests.at(-1)!.messages),
      ),
    ).toContain('Meeting Minutes Skill')
    await shellPage.evaluate(() => {
      const cb = window.workspaceServices.callbacks.at(-1)!
      cb.onToolCall({
        id: 'propose-from-workspace',
        name: 'propose_skill',
        input: {
          name: 'workspace-meeting-notes',
          description: 'Repeated meeting notes',
          body: 'Keep decisions and actions together.',
        },
      })
      cb.onDone()
    })
    await expect(panel.locator('.ai-skill-card')).toBeVisible()
    await expect
      .poll(() => shellPage.evaluate(() => window.workspaceServices.requests.length))
      .toBe(beforeSkills + 2)
    await shellPage.evaluate(() => {
      const cb = window.workspaceServices.callbacks.at(-1)!
      cb.onDelta('The skill proposal is ready to review')
      cb.onDone()
    })
    await expect
      .poll(() => shellPage.evaluate(() => window.workspaceServices.host.busy()))
      .toBe(false)
    const skillHistory = await shellPage.evaluate(async () => {
      const f = window.workspaceServices
      return f.client.load(f.host.conversations.state(f.selected()).origin)
    })
    expect(JSON.stringify(skillHistory)).not.toContain('Meeting Minutes Skill')
    expect(skillHistory.messages[0]!.skills).toEqual([
      expect.objectContaining({ id: 'builtin:meeting-minutes' }),
    ])
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
