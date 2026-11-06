import { test, expect } from '@playwright/test'
import { createServer } from 'node:http'
import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { buildDocx } from '../packages/docx-engine/tests/helpers/build-docx'
import { launchShell, waitForPageWithUrl, closeAndSaveVideo } from './helpers'

test('the default DOCX panel edits and restores review stores and refuses undo after a manual comment edit', async () => {
  test.setTimeout(120_000)
  const requests: { messages: { role: string; content: string }[] }[] = []
  const batches = [
    [
      ['add_comment', { blockIndex: 0, comment: 'WORKSPACE_COMMENT' }],
      ['insert_footnote', { blockIndex: 0, text: 'WORKSPACE_FOOTNOTE' }],
      ['insert_endnote', { blockIndex: 0, text: 'WORKSPACE_ENDNOTE' }],
      ['read_comments', {}],
      ['read_notes', {}],
    ],
    [
      ['resolve_comment', { id: '1' }],
      ['edit_note', { kind: 'footnote', id: '1', find: 'WORKSPACE', replace: 'EDITED' }],
      ['read_comments', {}],
      ['read_notes', {}],
    ],
    [
      ['read_comments', {}],
      ['read_notes', {}],
    ],
  ] as const
  const server = createServer(async (request, response) => {
    const chunks: Buffer[] = []
    for await (const chunk of request) chunks.push(Buffer.from(chunk))
    requests.push(JSON.parse(Buffer.concat(chunks).toString('utf8')))
    const turn = requests.length
    const delta =
      turn % 2
        ? {
            tool_calls: batches[(turn - 1) / 2].map(([name, input], index) => ({
              index,
              id: `review-${turn}-${index}`,
              type: 'function',
              function: { name, arguments: JSON.stringify(input) },
            })),
          }
        : { content: `REVIEW_DONE_${turn}` }
    response.writeHead(200, { 'Content-Type': 'text/event-stream' })
    response.end(
      `data: ${JSON.stringify({ choices: [{ delta, finish_reason: turn % 2 ? 'tool_calls' : 'stop' }] })}\n\ndata: [DONE]\n\n`,
    )
  })
  await new Promise<void>((done) => server.listen(0, '127.0.0.1', done))
  const address = server.address()
  if (!address || typeof address === 'string') throw new Error('No local endpoint')
  const scratch = await mkdtemp(join(tmpdir(), 'workspace-docs-review-'))
  const path = join(scratch, 'review.docx')
  const original = await buildDocx({
    bodyXml: '<w:p><w:r><w:t>Review fixture body</w:t></w:r></w:p>',
  })
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
    videoDir: 'workspace-docs-review',
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '/shell/out/renderer/index.html')
    const docs = await waitForPageWithUrl(launched.app, '://docs/')
    const editor = docs.locator('.ProseMirror').first()
    await expect(editor).toContainText('Review fixture body')
    await page.locator('.tab-split-btn[aria-pressed]').click()
    const panel = page.locator('.workspace-ai-surface')
    await panel.getByRole('button', { name: 'New conversation', exact: true }).click()
    expect(requests).toHaveLength(0)
    const send = async (text: string, turn: number) => {
      await panel.getByRole('textbox').fill(text)
      await panel.getByRole('button', { name: 'Send', exact: true }).click()
      await expect(panel).toContainText(`REVIEW_DONE_${turn}`)
    }
    const lastRead = (turn: number) =>
      requests[turn - 1].messages.filter((message) => message.role === 'tool').slice(-2)
    await send('Add a comment, footnote and endnote', 2)
    expect(lastRead(2)[0].content).toContain('WORKSPACE_COMMENT')
    expect(lastRead(2)[1].content).toContain('WORKSPACE_FOOTNOTE')
    expect(lastRead(2)[1].content).toContain('WORKSPACE_ENDNOTE')
    const beforeStoreOnly = await editor.textContent()
    await send('Resolve the comment and edit the footnote', 4)
    expect(lastRead(4)[0].content).toContain('[resolved]')
    expect(lastRead(4)[1].content).toContain('EDITED_FOOTNOTE')
    expect(await editor.textContent()).toBe(beforeStoreOnly)
    await panel.getByRole('button', { name: 'Undo this run', exact: true }).last().click()
    await expect(panel).toContainText('Reverted')
    await expect(editor.locator('.doc-comment-resolved')).toHaveCount(0)
    await send('Read the restored review lists', 6)
    expect(lastRead(6)[0].content).not.toContain('[resolved]')
    expect(lastRead(6)[1].content).toContain('WORKSPACE_FOOTNOTE')
    expect(lastRead(6)[1].content).not.toContain('EDITED_FOOTNOTE')
    expect(await readFile(path)).toEqual(Buffer.from(original))

    // Edit only the separate comment store through the production comments UI.
    await docs.locator('.ribbon-tab', { hasText: /^Review$/ }).click()
    await docs.getByRole('button', { name: 'Show Comments', exact: true }).click()
    const comments = docs.locator('.comments-pane')
    await comments.getByRole('button', { name: 'Edit comment', exact: true }).click()
    await comments.locator('textarea').fill('MANUAL_COMMENT_AFTER_RUN')
    await comments.getByRole('button', { name: 'Save', exact: true }).click()
    await expect(comments).toContainText('MANUAL_COMMENT_AFTER_RUN')
    await panel.getByRole('button', { name: 'Undo this run', exact: true }).click()
    await expect(
      panel.getByRole('button', { name: 'Could not restore. Check the document.' }),
    ).toBeDisabled()
    await expect(comments).toContainText('MANUAL_COMMENT_AFTER_RUN')

    await editor.click()
    await docs.keyboard.press(process.platform === 'darwin' ? 'Meta+s' : 'Control+s')
    await expect
      .poll(async () => {
        const zip = await JSZip.loadAsync(await readFile(path))
        return zip.file('word/comments.xml')?.async('string')
      })
      .toContain('MANUAL_COMMENT_AFTER_RUN')
    const saved = await JSZip.loadAsync(await readFile(path))
    expect(await saved.file('word/footnotes.xml')!.async('string')).toContain('WORKSPACE_FOOTNOTE')
    expect(await saved.file('word/endnotes.xml')!.async('string')).toContain('WORKSPACE_ENDNOTE')
    expect(requests).toHaveLength(6)
  } finally {
    await closeAndSaveVideo(launched, 'workspace-docs-review')
    server.closeAllConnections()
    await new Promise<void>((done) => server.close(() => done()))
  }
})
