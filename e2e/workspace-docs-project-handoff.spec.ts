import { test, expect } from '@playwright/test'
import { createServer } from 'node:http'
import { mkdtemp, readFile, readdir, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { buildDocx } from '../packages/docx-engine/tests/helpers/build-docx'
import { readDocxPartBase64 } from '@genoffice/docx-engine'
import type { ProjectFolderApi } from '@genoffice/project-store'
import { launchShell, waitForPageWithUrl, closeAndSaveVideo } from './helpers'

for (const outcome of ['success', 'cancel'] as const) {
  test(`a project template handoff keeps its context and ${outcome === 'success' ? 'saves its completed copy' : 'leaves a cancelled copy untitled'}`, async () => {
    test.setTimeout(120_000)
    const requests: unknown[] = []
    let release!: () => void
    const waiting = new Promise<void>((resolve) => {
      release = resolve
    })
    const server = createServer(async (request, response) => {
      const chunks: Buffer[] = []
      for await (const chunk of request) chunks.push(Buffer.from(chunk))
      requests.push(JSON.parse(Buffer.concat(chunks).toString('utf8')))
      if (requests.length === 1) await waiting
      const delta =
        requests.length === 1
          ? {
              tool_calls: [
                {
                  index: 0,
                  id: 'fill-copy',
                  type: 'function',
                  function: {
                    name: 'insert_content',
                    arguments: JSON.stringify({
                      afterBlockIndex: 0,
                      html: '<p>PROJECT_TEMPLATE_RESULT</p>',
                    }),
                  },
                },
              ],
            }
          : { content: 'PROJECT_HANDOFF_FINISHED' }
      response.writeHead(200, { 'Content-Type': 'text/event-stream' })
      response.end(
        `data: ${JSON.stringify({ choices: [{ delta, finish_reason: requests.length === 1 ? 'tool_calls' : 'stop' }] })}\n\ndata: [DONE]\n\n`,
      )
    })
    await new Promise<void>((resolve) => server.listen(0, '127.0.0.1', resolve))
    const address = server.address()
    if (!address || typeof address === 'string') throw new Error('No fixture endpoint')
    const scratch = await mkdtemp(join(tmpdir(), 'workspace-project-copy-'))
    const projectParent = await mkdtemp(join(tmpdir(), 'workspace-project-folder-'))
    const sourcePath = join(scratch, 'original.docx')
    const original = await buildDocx({
      bodyXml: '<w:p><w:r><w:t>ORIGINAL_TEMPLATE</w:t></w:r></w:p>',
    })
    await writeFile(sourcePath, original)
    const config = {
      apiKey: '',
      model: 'local-fixture',
      baseUrl: `http://127.0.0.1:${address.port}/v1`,
    }
    await writeFile(
      join(scratch, 'ai-settings.json'),
      JSON.stringify({
        provider: 'custom',
        providers: { custom: config },
        customEndpoint: 'fixture',
        customEndpoints: [
          {
            id: 'fixture',
            name: 'Local fixture',
            kind: 'local',
            ...config,
            modelProfiles: {
              'local-fixture': { contextLength: 100_000, tools: true, vision: false },
            },
          },
        ],
        subscriptionNoticeSeen: true,
        subscriptionNoticeVersion: 2,
      }),
    )
    const launched = await launchShell({
      userDataDir: scratch,
      onboardingSeen: true,
      settings: { defaultSaveDir: projectParent },
      openFile: sourcePath,
      recordVideo: false,
      videoDir: 'workspace-project-copy',
    })
    try {
      const page = await waitForPageWithUrl(launched.app, '/shell/out/renderer/index.html')
      const source = await waitForPageWithUrl(launched.app, '://docs/')
      await expect(source.locator('.ProseMirror').first()).toContainText('ORIGINAL_TEMPLATE')
      await page.locator('.tab-split-btn[aria-pressed]').click()
      const panel = page.locator('.workspace-ai-surface')
      await panel.getByRole('button', { name: 'New conversation', exact: true }).click()
      await panel.getByRole('textbox').fill('PRESERVE_OTHER_PROJECT_DRAFT')
      const conversation = panel.getByRole('combobox', { name: 'Conversation', exact: true })
      const previous = await conversation.inputValue()
      const project = await page.evaluate(async () => {
        const bridge = (window as unknown as { projectFolderApi: ProjectFolderApi })
          .projectFolderApi
        const made = await bridge.createProject({ name: 'Template project' })
        if (made.status !== 'created') throw new Error(`Project creation failed: ${made.status}`)
        await bridge.saveInstructions({
          projectId: made.project.id,
          text: 'PROJECT_CONTEXT_FOR_TEMPLATE',
        })
        return made.project
      })
      // Exercise a fresh untitled editor and the real one-use slot, without installing
      // a fixture skill in the person's home. Byte-for-byte template opening is covered
      // by skill-template-tabs.test.ts and template-copy-tabs.test.ts.
      await page.evaluate(() => window.aiOffice.newDoc())
      await expect
        .poll(
          () =>
            launched.app
              .context()
              .pages()
              .filter((p) => p.url().includes('://docs/')).length,
        )
        .toBe(2)
      const copy = launched.app
        .context()
        .pages()
        .find((p) => p !== source && p.url().includes('://docs/'))!
      await expect(copy.locator('.ProseMirror').first()).toBeVisible()
      await copy.evaluate(async (projectId) => {
        await (
          window as unknown as { projectFolderApi: ProjectFolderApi }
        ).projectFolderApi.setSelection({ projectId })
      }, project.id)
      await launched.app.evaluate(
        ({ ipcMain, webContents }, { projectId, copyUrl }) => {
          const skill = {
            id: 'app:project-handoff',
            name: 'project-handoff',
            source: 'app',
            title: 'Project template',
            description: 'Fill a copy',
            enabled: true,
            invoke: 'auto',
            owned: true,
            template: 'assets/form.docx',
            templateExists: true,
          }
          ipcMain.removeHandler('skills:list')
          ipcMain.handle('skills:list', () => [skill])
          ipcMain.removeHandler('skills:ai-read')
          ipcMain.handle('skills:ai-read', () => ({
            ok: true,
            skill: { ...skill, body: 'CURRENT_PROJECT_TEMPLATE_SKILL', files: [] },
          }))
          const contents = webContents.getAllWebContents().find((w) => w.getURL() === copyUrl)!
          const store = (
            globalThis as unknown as Record<symbol, { skillQueued?: Map<number, unknown> }>
          )[Symbol.for('genoffice.generate-handoff')]!
          store.skillQueued ??= new Map()
          store.skillQueued.set(contents.id, {
            id: `project-handoff-${projectId}`,
            mode: 'skill',
            target: 'docs',
            text: 'Fill the project template',
            chips: [{ id: skill.id, source: skill.source, title: skill.title }],
            createdAt: Date.now(),
          })
          for (const wc of webContents.getAllWebContents()) wc.send('skills:changed')
        },
        { projectId: project.id, copyUrl: copy.url() },
      )
      expect(
        await copy.evaluate(() =>
          (
            window as unknown as {
              aiOfficeUserSkills: { takeHandoff(editor: string): Promise<unknown> }
            }
          ).aiOfficeUserSkills.takeHandoff('docs'),
        ),
      ).toBeNull()
      await expect.poll(() => requests.length).toBe(1)
      expect(JSON.stringify(requests)).toContain('PROJECT_CONTEXT_FOR_TEMPLATE')
      expect(JSON.stringify(requests)).toContain('CURRENT_PROJECT_TEMPLATE_SKILL')
      const handedConversation = await conversation.inputValue()
      expect(handedConversation).not.toBe(previous)
      await conversation.selectOption(previous)
      await expect(panel.getByRole('textbox')).toHaveValue('PRESERVE_OTHER_PROJECT_DRAFT')
      await page.locator('.tab-item', { hasText: 'original.docx' }).click()
      if (outcome === 'cancel') {
        await conversation.selectOption(handedConversation)
        await panel.getByRole('button', { name: 'Stop', exact: true }).click()
        release()
        await expect(panel.getByRole('button', { name: 'Stop', exact: true })).toHaveCount(0)
        expect((await readdir(project.path)).filter((name) => name.endsWith('.docx'))).toEqual([])
        expect(requests).toHaveLength(1)
      } else {
        release()
        await expect
          .poll(
            async () =>
              (await readdir(project.path)).filter((name) => name.endsWith('.docx')).length,
          )
          .toBe(1)
        const savedName = (await readdir(project.path)).find((name) => name.endsWith('.docx'))!
        const saved = new Uint8Array(await readFile(join(project.path, savedName)))
        expect(
          Buffer.from((await readDocxPartBase64(saved, 'word/document.xml'))!, 'base64').toString(
            'utf8',
          ),
        ).toContain('PROJECT_TEMPLATE_RESULT')
        await expect(copy.locator('.ProseMirror').first()).toContainText('PROJECT_TEMPLATE_RESULT')
        await conversation.selectOption(handedConversation)
        await expect(panel).toContainText('PROJECT_HANDOFF_FINISHED')
        const chatDir = join(project.path, '.ieumoffice', 'chats')
        await expect
          .poll(
            async () =>
              (await readdir(chatDir).catch(() => [])).filter((name) => name.endsWith('.jsonl'))
                .length,
          )
          .toBe(1)
        const history = await readFile(
          join(
            chatDir,
            (await readdir(chatDir)).find((name) => name.endsWith('.jsonl'))!,
          ),
          'utf8',
        )
        expect(history).toContain('PROJECT_HANDOFF_FINISHED')
        expect(history).not.toContain('CURRENT_PROJECT_TEMPLATE_SKILL')
        expect(history).not.toContain('PROJECT_CONTEXT_FOR_TEMPLATE')
        expect(requests).toHaveLength(2)
      }
      expect(await readFile(sourcePath)).toEqual(Buffer.from(original))
      await conversation.selectOption(previous)
      await expect(panel.getByRole('textbox')).toHaveValue('PRESERVE_OTHER_PROJECT_DRAFT')
    } finally {
      release()
      await closeAndSaveVideo(launched, 'workspace-project-copy')
      await new Promise<void>((resolve) => {
        server.closeAllConnections()
        server.close(() => resolve())
      })
    }
  })
}
