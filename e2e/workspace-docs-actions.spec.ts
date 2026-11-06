import { test, expect } from '@playwright/test'
import { createServer, type ServerResponse } from 'node:http'
import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { buildDocx } from '../packages/docx-engine/tests/helpers/build-docx'
import { launchShell, waitForPageWithUrl, closeAndSaveVideo } from './helpers'

test('DOCX quick actions fill drafts, preserve the pinned selection, and keep delayed maps in the original conversation', async () => {
  test.setTimeout(120_000)
  const requests: { messages: { role: string; content: string }[] }[] = []
  let mapResponse: ServerResponse | null = null
  const reply = (response: ServerResponse, delta: object) => {
    response.writeHead(200, { 'Content-Type': 'text/event-stream' })
    response.end(`data: ${JSON.stringify({ choices: [{ delta }] })}\n\ndata: [DONE]\n\n`)
  }
  const tree = {
    label: 'PINNED_MAP_TOPIC',
    block: 0,
    children: [{ label: 'Map branch', block: 1 }],
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
    if (requests.length === 3) {
      mapResponse = response
      return
    }
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
    await editor.locator('p').first().selectText()
    await page.locator('.tab-split-btn[aria-pressed]').click()
    const panel = page.locator('.workspace-ai-surface')
    await panel.getByRole('button', { name: 'New conversation', exact: true }).click()
    const conversation = panel.getByRole('combobox', { name: 'Conversation', exact: true })
    const p = await conversation.inputValue()
    const target = panel.getByRole('combobox', { name: 'Edit target', exact: true })
    // A later visible selection and another focused tab must not replace the captured range.
    await editor.locator('p').nth(1).selectText()
    await page.evaluate((file) => window.aiOffice.openPath(file), other)
    await expect(target.getByRole('option', { name: /other\.docx/ })).toBeEnabled()
    await expect(target.locator('option:checked')).toContainText('pinned.docx')
    const input = panel.getByRole('textbox')
    for (const name of ['AI Summarize', 'AI Polish', 'AI Format', 'AI Mind Map']) {
      await panel.getByRole('button', { name, exact: true }).click()
      await expect(input).not.toHaveValue('')
      await expect(input).toBeFocused()
      expect(requests).toHaveLength(0)
    }
    await panel.getByRole('button', { name: 'AI Polish', exact: true }).click()
    await expect(input).toHaveValue(/pinned selection/)
    await panel.getByRole('button', { name: 'Send', exact: true }).click()
    await expect(panel).toContainText('ACTION_DONE_2')
    await expect(editor).toContainText('POLISHED_SELECTION')
    await expect(editor).toContainText('UNSELECTED_PARAGRAPH')
    await expect(editor).not.toContainText('ORIGINAL_SELECTION')
    expect(JSON.stringify(requests[0])).not.toContain('OTHER_DOCUMENT')
    await panel.getByRole('button', { name: 'Undo this run', exact: true }).click()
    await expect(editor).toContainText('ORIGINAL_SELECTION')
    await expect(editor).not.toContainText('POLISHED_SELECTION')

    await panel.getByRole('button', { name: 'AI Mind Map', exact: true }).click()
    expect(requests).toHaveLength(2)
    await panel.getByRole('button', { name: 'Send', exact: true }).click()
    await expect.poll(() => requests.length).toBe(3)
    await panel.getByRole('button', { name: 'New conversation', exact: true }).click()
    await input.fill('KEEP_OTHER_CONVERSATION')
    reply(mapResponse!, tool('generate_mindmap', { root: tree }))
    await expect.poll(() => requests.length).toBe(4)
    await expect(panel.locator('.ai-mindmap')).toHaveCount(0)
    await expect(input).toHaveValue('KEEP_OTHER_CONVERSATION')
    await conversation.selectOption(p)
    await expect(panel).toContainText('ACTION_DONE_4')
    const map = panel.locator('.ai-mindmap')
    await expect(map).toContainText('PINNED_MAP_TOPIC')
    await expect(map.getByRole('button', { name: 'Collapse branch', exact: true })).toBeVisible()
    await map.getByRole('button', { name: 'Collapse branch', exact: true }).click()
    await expect(map.getByRole('button', { name: 'Map branch', exact: true })).toHaveCount(0)
    await map.getByRole('button', { name: 'Expand branch', exact: true }).click()
    const svgPath = join(scratch, 'mindmap.svg')
    await launched.app.evaluate(({ session }, destination) => {
      session.defaultSession.once('will-download', (_event, item) => item.setSavePath(destination))
    }, svgPath)
    await map.getByRole('button', { name: 'Export SVG', exact: true }).click()
    await expect
      .poll(async () => readFile(svgPath, 'utf8').catch(() => ''))
      .toContain('PINNED_MAP_TOPIC')
    await map.getByRole('button', { name: 'PINNED_MAP_TOPIC', exact: true }).click()
    await expect(input).toHaveValue(/PINNED_MAP_TOPIC/)
    expect(requests).toHaveLength(4)
    await target.selectOption({
      label: (await target.getByRole('option', { name: /other\.docx/ }).textContent())!,
    })
    await input.fill('KEEP_TARGET_DRAFT')
    await expect(map.locator('.ai-mindmap-node[role="button"]')).toHaveCount(0)
    await map.locator('.ai-mindmap-node').first().click()
    await expect(input).toHaveValue('KEEP_TARGET_DRAFT')
    expect(requests).toHaveLength(4)
    expect(await readFile(path)).toEqual(Buffer.from(original))
    expect(await readFile(other)).toEqual(Buffer.from(otherBytes))
    await page.screenshot({ path: test.info().outputPath('workspace-docs-actions.png') })
  } finally {
    await closeAndSaveVideo(launched, 'workspace-docs-actions')
    server.closeAllConnections()
    await new Promise<void>((done) => server.close(() => done()))
  }
})
