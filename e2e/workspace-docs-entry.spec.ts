import { test, expect } from '@playwright/test'
import { createServer, type ServerResponse } from 'node:http'
import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { buildDocx } from '../packages/docx-engine/tests/helpers/build-docx'
import { launchShell, waitForPageWithUrl, closeAndSaveVideo } from './helpers'

test('Word ribbon and native menu enter the shared host while preserving older drafts and the clicked selection', async () => {
  test.setTimeout(120_000)
  const requests: { messages: { role: string; content: string }[] }[] = []
  const reply = (response: ServerResponse, delta: object) => {
    response.writeHead(200, { 'Content-Type': 'text/event-stream' })
    response.end(`data: ${JSON.stringify({ choices: [{ delta }] })}\n\ndata: [DONE]\n\n`)
  }
  const tool = (name: string, input: object) => ({
    tool_calls: [
      {
        index: 0,
        id: `action-${requests.length}`,
        type: 'function',
        function: { name, arguments: JSON.stringify(input) },
      },
    ],
  })
  const server = createServer(async (request, response) => {
    const chunks: Buffer[] = []
    for await (const chunk of request) chunks.push(Buffer.from(chunk))
    requests.push(JSON.parse(Buffer.concat(chunks).toString('utf8')))
    reply(
      response,
      requests.length === 1
        ? tool('replace_selection', { html: 'POLISHED_SELECTION' })
        : { content: `ACTION_DONE_${requests.length}` },
    )
  })
  await new Promise<void>((done) => server.listen(0, '127.0.0.1', done))
  const address = server.address()
  if (!address || typeof address === 'string') throw new Error('No local endpoint')
  const scratch = await mkdtemp(join(tmpdir(), 'workspace-docs-actions-'))
  const path = join(scratch, 'pinned.docx')
  const other = join(scratch, 'other.docx')
  const original = await buildDocx({
    bodyXml:
      '<w:p><w:r><w:t>ORIGINAL_SELECTION</w:t></w:r></w:p><w:p><w:r><w:t>UNSELECTED_PARAGRAPH</w:t></w:r></w:p>',
  })
  const otherBytes = await buildDocx({ bodyXml: '<w:p><w:r><w:t>OTHER_DOCUMENT</w:t></w:r></w:p>' })
  await writeFile(path, original)
  await writeFile(other, otherBytes)
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
    videoDir: 'workspace-docs-actions',
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '/shell/out/renderer/index.html')
    const docs = await waitForPageWithUrl(launched.app, '://docs/')
    const editor = docs.locator('.ProseMirror').first()
    await expect(editor).toContainText('ORIGINAL_SELECTION')
    await page.locator('.tab-split-btn[aria-pressed]').click()
    const panel = page.locator('.workspace-ai-surface')
    await panel.getByRole('button', { name: 'New conversation', exact: true }).click()
    const conversation = panel.getByRole('combobox', { name: 'Conversation', exact: true })
    const previous = await conversation.inputValue()
    await panel.getByRole('textbox').fill('KEEP_OLDER_DRAFT')
    await page.locator('.tab-split-btn[aria-pressed]').click()
    await editor.locator('p').first().selectText()
    await docs.locator('button[data-tip="AI Polish"]').click()
    await expect(panel).toContainText('ACTION_DONE_2')
    await expect(editor).toContainText('POLISHED_SELECTION')
    await expect(editor).toContainText('UNSELECTED_PARAGRAPH')
    expect(await conversation.inputValue()).not.toBe(previous)
    await expect(
      panel.getByRole('combobox', { name: 'Edit target', exact: true }).locator('option:checked'),
    ).toContainText('pinned.docx')
    await expect(docs.locator('.ai-dock')).toHaveClass(/collapsed/)
    expect(requests).toHaveLength(2)
    expect(JSON.stringify(requests[0])).not.toContain('OTHER_DOCUMENT')
    await conversation.selectOption(previous)
    await expect(panel.getByRole('textbox')).toHaveValue('KEEP_OLDER_DRAFT')
    const toggle = () =>
      launched.app.evaluate(({ Menu }) => {
        Menu.getApplicationMenu()?.getMenuItemById('docs-menu-ai-sidebar')?.click()
      })
    await toggle()
    await expect(panel).toHaveCount(0)
    await toggle()
    await expect(panel.getByRole('textbox')).toHaveValue('KEEP_OLDER_DRAFT')
    expect(await conversation.inputValue()).toBe(previous)
    await editor.locator('p').first().selectText()
    await editor.locator('p').first().click({ button: 'right' })
    await docs
      .locator('.ctx-item')
      .filter({ hasText: /^Synonyms$/ })
      .click()
    await expect(panel).toContainText('ACTION_DONE_3')
    expect(await conversation.inputValue()).not.toBe(previous)
    await conversation.selectOption(previous)
    await expect(panel.getByRole('textbox')).toHaveValue('KEEP_OLDER_DRAFT')
    expect(requests).toHaveLength(3)
    expect(await readFile(path)).toEqual(Buffer.from(original))
    expect(await readFile(other)).toEqual(Buffer.from(otherBytes))
    await page.screenshot({ path: test.info().outputPath('workspace-docs-entry.png') })
  } finally {
    await closeAndSaveVideo(launched, 'workspace-docs-entry')
    server.closeAllConnections()
    await new Promise<void>((done) => server.close(() => done()))
  }
})
