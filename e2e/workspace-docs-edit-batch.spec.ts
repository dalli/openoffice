import { test, expect } from '@playwright/test'
import { createServer, type ServerResponse } from 'node:http'
import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { buildDocx } from '../packages/docx-engine/tests/helpers/build-docx'
import { launchShell, waitForPageWithUrl, closeAndSaveVideo } from './helpers'

test('Word queued passages execute in their shared conversation and undo together after a tab switch', async () => {
  test.setTimeout(120_000)
  const requests: object[] = []
  let release!: () => void
  const held = new Promise<void>((done) => {
    release = done
  })
  const server = createServer(async (request, response: ServerResponse) => {
    const chunks: Buffer[] = []
    for await (const chunk of request) chunks.push(Buffer.from(chunk))
    requests.push(JSON.parse(Buffer.concat(chunks).toString('utf8')))
    const first = requests.length === 1
    if (first) await held
    const delta = first
      ? {
          tool_calls: [2, 0].map((index, order) => ({
            index: order,
            id: `batch-${index}`,
            type: 'function',
            function: {
              name: 'replace_blocks',
              arguments: JSON.stringify({
                startBlockIndex: index,
                endBlockIndex: index,
                html: `<p>EDITED_${index}</p>`,
              }),
            },
          })),
        }
      : { content: 'BATCH_DONE' }
    response.writeHead(200, { 'Content-Type': 'text/event-stream' })
    response.end(`data: ${JSON.stringify({ choices: [{ delta }] })}\n\ndata: [DONE]\n\n`)
  })
  await new Promise<void>((done) => server.listen(0, '127.0.0.1', done))
  const address = server.address()
  if (!address || typeof address === 'string') throw new Error('No local endpoint')
  const scratch = await mkdtemp(join(tmpdir(), 'workspace-docs-batch-'))
  const path = join(scratch, 'batch.docx')
  const other = join(scratch, 'other.docx')
  const original = await buildDocx({
    bodyXml: ['FIRST_PASSAGE', 'KEEP_MIDDLE', 'LAST_PASSAGE']
      .map((text) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`)
      .join(''),
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
    videoDir: 'workspace-docs-batch',
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '/shell/out/renderer/index.html')
    const docs = await waitForPageWithUrl(launched.app, '://docs/')
    const editor = docs.locator('.ProseMirror').first()
    await expect(editor).toContainText('FIRST_PASSAGE')
    await page.locator('.tab-split-btn[aria-pressed]').click()
    const panel = page.locator('.workspace-ai-surface')
    await panel.getByRole('button', { name: 'New conversation', exact: true }).click()
    const conversation = panel.getByRole('combobox', { name: 'Conversation', exact: true })
    const previous = await conversation.inputValue()
    await panel.getByRole('textbox').fill('KEEP_PREVIOUS_DRAFT')
    await page.locator('.tab-split-btn[aria-pressed]').click()
    for (const [index, instruction] of [
      [0, 'Polish first passage'],
      [2, 'Shorten last passage'],
    ] as const) {
      await editor.locator('p').nth(index).selectText()
      await docs.locator('.ai-ask-trigger').click()
      await docs.locator('.ai-ask-pop-input').fill(instruction)
      await docs.getByRole('button', { name: 'Add to queue', exact: true }).click()
    }
    const queue = panel.getByRole('group', { name: 'Queued edits', exact: true })
    await expect(queue).toContainText('2/10')
    expect(requests).toHaveLength(0)
    await queue.locator('.ai-queue-text').first().click()
    await expect
      .poll(() => docs.evaluate(() => window.getSelection()?.toString()))
      .toBe('FIRST_PASSAGE')
    await queue.getByRole('button', { name: 'Edit', exact: true }).first().click()
    await queue
      .getByRole('textbox', { name: 'Edit', exact: true })
      .fill('Polish first passage carefully')
    await queue.getByRole('textbox', { name: 'Edit', exact: true }).press('Enter')
    await expect(queue).toContainText('Polish first passage carefully')
    await queue.getByRole('button', { name: 'Discard', exact: true }).click()
    await expect(queue).toContainText('Discard all 2 queued edits?')
    await queue.getByRole('button', { name: 'Cancel', exact: true }).click()
    // The editor stays responsible for anchors; adding another annotation only updates this card.
    await editor.locator('p').nth(1).selectText()
    await docs.locator('.ai-ask-trigger').click()
    await docs.locator('.ai-ask-pop-input').fill('REMOVE_THIS_REQUEST')
    await docs.getByRole('button', { name: 'Add to queue', exact: true }).click()
    await expect(queue).toContainText('3/10')
    await queue
      .locator('.ai-queue-row')
      .filter({ hasText: 'REMOVE_THIS_REQUEST' })
      .getByRole('button', { name: 'Delete', exact: true })
      .click()
    await expect(queue).toContainText('2/10')
    await expect(queue).not.toContainText('REMOVE_THIS_REQUEST')
    expect(requests).toHaveLength(0)
    await page.screenshot({ path: test.info().outputPath('workspace-docs-queue-card.png') })
    await page.evaluate((file) => window.aiOffice.openPath(file), other)
    await expect(queue).toContainText('Polish first passage carefully')
    await queue.getByRole('button', { name: 'Send 2 edits', exact: true }).click()
    await expect.poll(() => requests.length).toBe(1)
    const batchConversation = await conversation.inputValue()
    expect(batchConversation).not.toBe(previous)
    await expect(panel).toContainText('Batch of 2 edits')
    await expect(docs.locator('.ai-dock')).toHaveClass(/collapsed/)
    await expect(docs.locator('.ai-queue-anchor')).toHaveCount(0)
    expect(JSON.stringify(requests[0])).toContain('Polish first passage carefully')
    expect(JSON.stringify(requests[0])).not.toContain('REMOVE_THIS_REQUEST')
    expect(JSON.stringify(requests[0])).toContain('Shorten last passage')
    expect(JSON.stringify(requests[0])).not.toContain('OTHER_DOCUMENT')
    await conversation.selectOption(previous)
    await expect(panel.getByRole('textbox')).toHaveValue('KEEP_PREVIOUS_DRAFT')
    release()
    await expect(editor).toContainText('EDITED_0')
    await expect(editor).toContainText('EDITED_2')
    await expect(editor).toContainText('KEEP_MIDDLE')
    await conversation.selectOption(batchConversation)
    await expect(panel).toContainText('BATCH_DONE')
    await panel.getByRole('button', { name: 'Undo this run', exact: true }).click()
    await expect(editor).toContainText('FIRST_PASSAGE')
    await expect(editor).toContainText('LAST_PASSAGE')
    await expect(editor).not.toContainText('EDITED_')
    expect(requests).toHaveLength(2)
    expect(await readFile(path)).toEqual(Buffer.from(original))
    expect(await readFile(other)).toEqual(Buffer.from(otherBytes))
    await page.screenshot({ path: test.info().outputPath('workspace-docs-batch.png') })
  } finally {
    release()
    await closeAndSaveVideo(launched, 'workspace-docs-batch')
    server.closeAllConnections()
    await new Promise<void>((done) => server.close(() => done()))
  }
})
