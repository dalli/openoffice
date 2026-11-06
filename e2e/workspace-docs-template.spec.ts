import { test, expect } from '@playwright/test'
import { createServer } from 'node:http'
import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { buildDocx } from '../packages/docx-engine/tests/helpers/build-docx'
import { launchShell, waitForPageWithUrl, closeAndSaveVideo } from './helpers'

test('the default panel saves the proposal target as a template through the existing card bridge', async () => {
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
              ['insert_content', { afterBlockIndex: 0, html: '<p>UNSAVED_TEMPLATE_CONTENT</p>' }],
              [
                'propose_skill',
                {
                  name: 'workspace-template-fixture',
                  description: 'Reuse this document',
                  body: 'Fill in the document.',
                  wantsTemplate: true,
                },
              ],
            ].map(([name, input], index) => ({
              index,
              id: `template-${index}`,
              type: 'function',
              function: { name, arguments: JSON.stringify(input) },
            })),
          }
        : { content: 'TEMPLATE_READY' }
    response.writeHead(200, { 'Content-Type': 'text/event-stream' })
    response.end(
      `data: ${JSON.stringify({ choices: [{ delta, finish_reason: requests.length === 1 ? 'tool_calls' : 'stop' }] })}\n\ndata: [DONE]\n\n`,
    )
  })
  await new Promise<void>((done) => server.listen(0, '127.0.0.1', done))
  const address = server.address()
  if (!address || typeof address === 'string') throw new Error('No local endpoint')
  const scratch = await mkdtemp(join(tmpdir(), 'workspace-docs-template-'))
  const path = join(scratch, 'source.docx')
  const copied = join(scratch, 'template.docx')
  const original = await buildDocx({ bodyXml: '<w:p><w:r><w:t>Template source</w:t></w:r></w:p>' })
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
    videoDir: 'workspace-docs-template',
  })
  try {
    // Exercise the real card/preload and native template export, with a scratch-only save sink.
    // Never write test skills into the person's actual skill directory.
    await launched.app.evaluate(async ({ ipcMain }, copyPath) => {
      const { writeFileSync } = process.getBuiltinModule('fs')
      ipcMain.removeHandler('skills:save-from-card')
      ipcMain.handle('skills:save-from-card', async (_event, input) => {
        if (!input?.template || input.template.fileName !== 'source.docx')
          return { ok: false, reason: 'failed' }
        writeFileSync(copyPath, input.template.bytes)
        return { ok: true, id: 'app:workspace-template-fixture' }
      })
    }, copied)
    const page = await waitForPageWithUrl(launched.app, '/shell/out/renderer/index.html')
    const docs = await waitForPageWithUrl(launched.app, '://docs/')
    await expect(docs.locator('.ProseMirror').first()).toContainText('Template source')
    await page.locator('.tab-split-btn[aria-pressed]').click()
    const panel = page.locator('.workspace-ai-surface')
    await panel.getByRole('button', { name: 'New conversation', exact: true }).click()
    await panel.getByRole('textbox').fill('Add text and propose a reusable template')
    await panel.getByRole('button', { name: 'Send', exact: true }).click()
    await expect(panel).toContainText('TEMPLATE_READY')
    await expect(docs.locator('.ProseMirror').first()).toContainText('UNSAVED_TEMPLATE_CONTENT')
    const save = panel.getByRole('button', { name: 'Save the document as it is', exact: true })
    await expect(save).toBeEnabled()
    await save.click()
    await expect(panel).toContainText('Saved')
    await expect
      .poll(async () => {
        try {
          return await (
            await JSZip.loadAsync(await readFile(copied))
          )
            .file('word/document.xml')!
            .async('string')
        } catch {
          return ''
        }
      })
      .toContain('UNSAVED_TEMPLATE_CONTENT')
    expect(await readFile(path)).toEqual(Buffer.from(original))
    expect(requests).toHaveLength(2)
  } finally {
    await closeAndSaveVideo(launched, 'workspace-docs-template')
    server.closeAllConnections()
    await new Promise<void>((done) => server.close(() => done()))
  }
})
