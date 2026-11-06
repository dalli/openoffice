import { test, expect } from '@playwright/test'
import { createServer } from 'node:http'
import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { buildDocx } from '../packages/docx-engine/tests/helpers/build-docx'
import { launchShell, waitForPageWithUrl, closeAndSaveVideo } from './helpers'

test('a Word template handoff runs once in a new shared conversation and preserves the other draft', async () => {
  test.setTimeout(120_000)
  const requests: unknown[] = []
  const server = createServer(async (request, response) => {
    const chunks: Buffer[] = []
    for await (const chunk of request) chunks.push(Buffer.from(chunk))
    requests.push(JSON.parse(Buffer.concat(chunks).toString('utf8')))
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
                    html: '<p>FILLED_TEMPLATE_COPY</p>',
                  }),
                },
              },
            ],
          }
        : { content: 'HANDOFF_FINISHED' }
    response.writeHead(200, { 'Content-Type': 'text/event-stream' })
    response.end(
      `data: ${JSON.stringify({ choices: [{ delta, finish_reason: requests.length === 1 ? 'tool_calls' : 'stop' }] })}\n\ndata: [DONE]\n\n`,
    )
  })
  await new Promise<void>((done) => server.listen(0, '127.0.0.1', done))
  const address = server.address()
  if (!address || typeof address === 'string') throw new Error('No fixture endpoint')
  const scratch = await mkdtemp(join(tmpdir(), 'workspace-handoff-'))
  const path = join(scratch, 'copy.docx')
  const original = await buildDocx({ bodyXml: '<w:p><w:r><w:t>Template copy</w:t></w:r></w:p>' })
  await writeFile(path, original)
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
    openFile: path,
    recordVideo: false,
    videoDir: 'workspace-docs-handoff',
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '/shell/out/renderer/index.html')
    const docs = await waitForPageWithUrl(launched.app, '://docs/')
    await expect(docs.locator('.ProseMirror').first()).toContainText('Template copy')
    await page.locator('.tab-split-btn[aria-pressed]').click()
    const panel = page.locator('.workspace-ai-surface')
    await panel.getByRole('button', { name: 'New conversation', exact: true }).click()
    await panel.getByRole('textbox').fill('PRESERVE_OTHER_DRAFT')
    const conversation = panel.getByRole('combobox', { name: 'Conversation', exact: true })
    const previous = await conversation.inputValue()
    // Seed the real main-owned, one-use slot. Skill reads are scratch-only fixtures;
    // no test writes or imports a skill into the person's home directory.
    await launched.app.evaluate(({ ipcMain, webContents }) => {
      const skill = {
        id: 'app:handoff-fixture',
        name: 'handoff-fixture',
        source: 'app',
        title: 'Handoff fixture',
        description: 'Fill this template',
        enabled: true,
        invoke: 'auto',
        owned: true,
        template: 'assets/form.docx',
        templateExists: true,
      }
      ipcMain.removeHandler('skills:list')
      ipcMain.handle('skills:list', () => [skill])
      ipcMain.removeHandler('skills:ai-read')
      ipcMain.handle('skills:ai-read', (_event, id) =>
        id === skill.id
          ? { ok: true, skill: { ...skill, body: 'CURRENT_HANDOFF_SKILL_BODY', files: [] } }
          : { ok: false, reason: 'missing' },
      )
      ipcMain.removeHandler('skills:open-template')
      ipcMain.handle('skills:open-template', () => {
        throw new Error('A handed copy must not open another copy')
      })
      const contents = webContents.getAllWebContents().find((w) => w.getURL().includes('://docs/'))!
      const holder = globalThis as unknown as Record<symbol, { skillQueued?: Map<number, unknown> }>
      const store = holder[Symbol.for('genoffice.generate-handoff')]!
      store.skillQueued ??= new Map()
      store.skillQueued.set(contents.id, {
        id: 'real-handoff-slot',
        mode: 'skill',
        target: 'docs',
        text: 'Fill the copied template',
        chips: [{ id: skill.id, source: skill.source, title: skill.title }],
        createdAt: Date.now(),
      })
      for (const wc of webContents.getAllWebContents()) wc.send('skills:changed')
    })
    const result = await docs.evaluate(async () => {
      const bridge = (
        window as unknown as {
          aiOfficeUserSkills: { takeHandoff(editor: string): Promise<unknown> }
        }
      ).aiOfficeUserSkills
      return bridge.takeHandoff('docs')
    })
    expect(result).toBeNull()
    await expect(panel).toContainText('HANDOFF_FINISHED')
    await expect(docs.locator('.ProseMirror').first()).toContainText('FILLED_TEMPLATE_COPY')
    expect(requests).toHaveLength(2)
    expect(JSON.stringify(requests)).toContain('CURRENT_HANDOFF_SKILL_BODY')
    expect(
      await docs.evaluate(() =>
        (
          window as unknown as {
            aiOfficeUserSkills: { takeHandoff(editor: string): Promise<unknown> }
          }
        ).aiOfficeUserSkills.takeHandoff('docs'),
      ),
    ).toBeNull()
    await panel.getByRole('button', { name: 'Undo this run', exact: true }).click()
    await expect(docs.locator('.ProseMirror').first()).not.toContainText('FILLED_TEMPLATE_COPY')
    await conversation.selectOption(previous)
    await expect(panel.getByRole('textbox')).toHaveValue('PRESERVE_OTHER_DRAFT')
    // Also exercise the mounted editor hook, rather than only an explicit bridge take.
    await launched.app.evaluate(({ webContents }) => {
      const contents = webContents.getAllWebContents().find((w) => w.getURL().includes('://docs/'))!
      const holder = globalThis as unknown as Record<symbol, { skillQueued: Map<number, unknown> }>
      holder[Symbol.for('genoffice.generate-handoff')]!.skillQueued.set(contents.id, {
        id: 'mounted-handoff-slot',
        mode: 'skill',
        target: 'docs',
        text: 'MOUNTED_COPY_REQUEST',
        chips: [{ id: 'app:handoff-fixture', source: 'app', title: 'Handoff fixture' }],
        createdAt: Date.now(),
      })
    })
    await docs.reload()
    await expect(panel).toContainText('MOUNTED_COPY_REQUEST')
    await expect(panel).toContainText('HANDOFF_FINISHED')
    await expect(conversation).not.toHaveValue(previous)
    expect(requests).toHaveLength(3)
    await conversation.selectOption(previous)
    await expect(panel.getByRole('textbox')).toHaveValue('PRESERVE_OTHER_DRAFT')
    expect(await readFile(path)).toEqual(Buffer.from(original))
    expect(requests).toHaveLength(3)
  } finally {
    await closeAndSaveVideo(launched, 'workspace-docs-handoff')
    server.closeAllConnections()
    await new Promise<void>((done) => server.close(() => done()))
  }
})
