import { test, expect } from '@playwright/test'
import { createServer, type ServerResponse } from 'node:http'
import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { buildDocx } from '../packages/docx-engine/tests/helpers/build-docx'
import { launchShell, waitForPageWithUrl, closeAndSaveVideo } from './helpers'

test('the shared writer streams to its pinned document, isolates decisions, restores the run and discards stopped drafts', async () => {
  test.setTimeout(120_000)
  const writers: ServerResponse[] = []
  const requests: { tools?: unknown[]; messages: { role: string; content: unknown }[] }[] = []
  const delta = (response: ServerResponse, value: object) =>
    response.write(`data: ${JSON.stringify({ choices: [{ delta: value }] })}\n\n`)
  const end = (response: ServerResponse, partial = false) =>
    response.end(
      `data: ${JSON.stringify({ choices: [{ delta: {}, finish_reason: partial ? 'length' : 'stop' }] })}\n\ndata: [DONE]\n\n`,
    )
  const server = createServer(async (request, response) => {
    const body: Buffer[] = []
    for await (const chunk of request) body.push(Buffer.from(chunk))
    const data = JSON.parse(Buffer.concat(body).toString('utf8')) as (typeof requests)[number]
    requests.push(data)
    response.writeHead(200, { 'Content-Type': 'text/event-stream' })
    if (!data.tools?.length) {
      writers.push(response)
      delta(response, { content: `<p>WRITER_DRAFT_${writers.length}</p>` })
      if (writers.length === 2 || writers.length === 3) end(response, true)
    } else if (data.messages.at(-1)?.role === 'tool') {
      delta(response, { content: `WRITER_DONE_${writers.length}` })
      end(response)
    } else {
      delta(response, {
        tool_calls: [
          {
            index: 0,
            id: `writer-${requests.length}`,
            type: 'function',
            function: {
              name: 'write_document',
              arguments: JSON.stringify({
                plan: 'Write a factual chapter from the supplied fixture.',
                afterBlockIndex: 1,
              }),
            },
          },
        ],
      })
      end(response)
    }
  })
  await new Promise<void>((done) => server.listen(0, '127.0.0.1', done))
  const address = server.address()
  if (!address || typeof address === 'string') throw new Error('No local fixture endpoint')
  const scratch = await mkdtemp(join(tmpdir(), 'workspace-writer-'))
  const path = join(scratch, 'writer.docx')
  const other = join(scratch, 'other.docx')
  const original = await buildDocx({
    bodyXml:
      '<w:p><w:r><w:t>ORIGINAL_ONE</w:t></w:r></w:p><w:p><w:r><w:t>ORIGINAL_TWO</w:t></w:r></w:p>',
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
    videoDir: 'workspace-docs-writer',
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '/shell/out/renderer/index.html')
    const docs = await waitForPageWithUrl(launched.app, '://docs/')
    const editor = docs.locator('.ProseMirror').first()
    await expect(editor).toContainText('ORIGINAL_ONE')
    await page.locator('.tab-split-btn[aria-pressed]').click()
    const panel = page.locator('.workspace-ai-surface')
    await panel.getByRole('button', { name: 'New conversation', exact: true }).click()
    const conversation = panel.getByRole('combobox', { name: 'Conversation', exact: true })
    const p = await conversation.inputValue()
    const input = panel.getByRole('textbox')
    await input.fill('PINNED_WRITER_REQUEST')
    await panel.getByRole('button', { name: 'Send', exact: true }).click()
    await expect(editor).toContainText('WRITER_DRAFT_1')
    await docs.screenshot({ path: test.info().outputPath('writer-draft-document.png') })
    await expect(panel.getByRole('group', { name: /Writing the document/ })).toBeVisible()
    await page.evaluate((file) => window.aiOffice.openPath(file), other)
    await panel.getByRole('button', { name: 'New conversation', exact: true }).click()
    await input.fill('OTHER_DRAFT')
    delta(writers[0]!, { content: '<p>COMPLETED_CHAPTER</p>' })
    end(writers[0]!)
    await expect(editor).toContainText('COMPLETED_CHAPTER')
    await expect(input).toHaveValue('OTHER_DRAFT')
    await expect(panel).not.toContainText('WRITER_DONE_1')
    await conversation.selectOption(p)
    await expect(panel).toContainText('WRITER_DONE_1')
    await panel.getByRole('button', { name: 'Undo this run', exact: true }).click()
    await expect(editor).not.toContainText('WRITER_DRAFT_1')
    await expect(editor).not.toContainText('COMPLETED_CHAPTER')
    await expect(editor).toContainText('ORIGINAL_TWO')

    for (const [number, keep] of [
      [2, true],
      [3, false],
    ] as const) {
      await input.fill(`WRITE_PARTIAL_${number}`)
      await panel.getByRole('button', { name: 'Send', exact: true }).click()
      const choice = panel.getByRole('group', {
        name: 'Only part of the content was written',
        exact: true,
      })
      await expect(choice).toBeVisible()
      if (number === 2) {
        await choice.screenshot({ path: test.info().outputPath('writer-partial-choice.png') })
        await panel.getByRole('button', { name: 'New conversation', exact: true }).click()
        await input.fill('KEEP_PARTIAL_DECISION_SEPARATE')
        await expect(
          panel.getByRole('group', { name: 'Only part of the content was written', exact: true }),
        ).toHaveCount(0)
        await conversation.selectOption(p)
        await expect(choice).toBeVisible()
      }
      await expect(editor).toContainText(`WRITER_DRAFT_${number}`)
      await choice.getByRole('button', { name: keep ? 'Keep' : 'Discard', exact: true }).click()
      await expect(panel).toContainText(`WRITER_DONE_${number}`)
      if (keep) {
        await expect(editor).toContainText(`WRITER_DRAFT_${number}`)
        await panel.getByRole('button', { name: 'Undo this run', exact: true }).click()
      }
      await expect(editor).not.toContainText(`WRITER_DRAFT_${number}`)
    }
    await input.fill('STOP_WRITER_ONLY')
    await panel.getByRole('button', { name: 'Send', exact: true }).click()
    await expect(editor).toContainText('WRITER_DRAFT_4')
    await panel
      .getByRole('group', { name: /Writing the document/ })
      .getByRole('button', { name: 'Stop', exact: true })
      .click()
    const choice = panel.getByRole('group', {
      name: 'Only part of the content was written',
      exact: true,
    })
    await expect(choice).toBeVisible()
    await choice.getByRole('button', { name: 'Discard', exact: true }).click()
    await expect(editor).not.toContainText('WRITER_DRAFT_4')
    await expect(panel).toContainText('WRITER_DONE_4')
    await input.fill('CANCEL_WHOLE_RUN')
    await panel.getByRole('button', { name: 'Send', exact: true }).click()
    await expect(editor).toContainText('WRITER_DRAFT_5')
    await panel.locator('.ai-stop-btn').click()
    await expect(editor).not.toContainText('WRITER_DRAFT_5')
    await expect(
      panel.getByRole('group', { name: 'Only part of the content was written', exact: true }),
    ).toHaveCount(0)
    expect(await readFile(path)).toEqual(Buffer.from(original))
    expect(await readFile(other)).toEqual(Buffer.from(otherBytes))
    expect(requests.filter((r) => !r.tools?.length)).toHaveLength(5)
    expect(JSON.stringify(requests.filter((r) => !r.tools?.length))).not.toContain('OTHER_DOCUMENT')
    await page.screenshot({ path: test.info().outputPath('workspace-docs-writer.png') })
  } finally {
    await closeAndSaveVideo(launched, 'workspace-docs-writer')
    server.closeAllConnections()
    await new Promise<void>((done) => server.close(() => done()))
  }
})
