import { test, expect } from '@playwright/test'
import { createServer } from 'node:http'
import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'
import { buildDocx } from '../packages/docx-engine/tests/helpers/build-docx'
import { launchShell, waitForPageWithUrl, closeAndSaveVideo } from './helpers'

test('the default DOCX panel restores layout stores and inserts sections without automatically saving the file', async () => {
  test.setTimeout(120_000)
  const requests: { messages: { role: string; content: string }[] }[] = []
  const batches = [
    [
      ['set_header_footer', { kind: 'header', text: 'WORKSPACE_HEADER' }],
      ['set_header_footer', { kind: 'footer', text: 'WORKSPACE_FOOTER' }],
      ['set_page_setup', { orientation: 'landscape', margins: { left: '2cm' } }],
      ['define_style', { styleId: 'WorkspaceStyle', basedOn: 'Normal', run: { fontSize: 18 } }],
      [
        'apply_ops',
        { ops: [{ op: 'applyStyle', target: { blockIndexes: [0] }, styleId: 'WorkspaceStyle' }] },
      ],
      ['set_watermark', { text: 'WORKSPACE_DRAFT' }],
      ['get_document_context', {}],
      ['list_styles', {}],
    ],
    [
      ['insert_section_break', { afterBlockIndex: 0, type: 'nextPage' }],
      ['set_page_setup', { section: 1, orientation: 'landscape' }],
      ['set_header_footer', { kind: 'footer', text: 'SAVED_FOOTER' }],
      ['define_style', { styleId: 'SavedStyle', basedOn: 'Normal', run: { fontSize: 19 } }],
      [
        'apply_ops',
        { ops: [{ op: 'applyStyle', target: { blockIndexes: [2] }, styleId: 'SavedStyle' }] },
      ],
      ['get_document_context', {}],
      ['list_styles', {}],
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
            tool_calls: batches[Math.floor((turn - 1) / 2)].map(([name, input], index) => ({
              index,
              id: `layout-${turn}-${index}`,
              type: 'function',
              function: { name, arguments: JSON.stringify(input) },
            })),
          }
        : { content: `LAYOUT_DONE_${turn}` }
    response.writeHead(200, { 'Content-Type': 'text/event-stream' })
    response.end(
      `data: ${JSON.stringify({ choices: [{ delta, finish_reason: turn % 2 ? 'tool_calls' : 'stop' }] })}\n\ndata: [DONE]\n\n`,
    )
  })
  await new Promise<void>((done) => server.listen(0, '127.0.0.1', done))
  const address = server.address()
  if (!address || typeof address === 'string') throw new Error('No local endpoint')
  const scratch = await mkdtemp(join(tmpdir(), 'workspace-docs-layout-'))
  const path = join(scratch, 'layout.docx')
  const original = await buildDocx({
    bodyXml:
      '<w:p><w:r><w:t>Layout fixture body</w:t></w:r></w:p><w:p><w:r><w:t>Second layout paragraph</w:t></w:r></w:p>',
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
    videoDir: 'workspace-docs-layout',
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '/shell/out/renderer/index.html')
    const docs = await waitForPageWithUrl(launched.app, '://docs/')
    const editor = docs.locator('.ProseMirror').first()
    await expect(editor).toContainText('Layout fixture body')
    await page.locator('.tab-split-btn[aria-pressed]').click()
    const panel = page.locator('.workspace-ai-surface')
    await panel.getByRole('button', { name: 'New conversation', exact: true }).click()
    expect(requests).toHaveLength(0)
    const send = async (text: string, turn: number) => {
      await panel.getByRole('textbox').fill(text)
      await panel.getByRole('button', { name: 'Send', exact: true }).click()
      await expect(panel).toContainText(`LAYOUT_DONE_${turn}`)
      const results = requests[turn - 1].messages.filter((m) => m.role === 'tool')
      expect(results.map((m) => m.content).join('\n')).not.toMatch(
        /not available|does not exist|Error:/,
      )
      return results
    }
    const results = await send('Change the header, footer, page layout, styles and watermark', 2)
    expect(results.at(-2)!.content).toContain('WORKSPACE_HEADER')
    expect(results.at(-2)!.content).toContain('WORKSPACE_FOOTER')
    expect(results.at(-1)!.content).toContain('WorkspaceStyle')
    await expect(docs.locator('.page-hf-header').first()).toContainText('WORKSPACE_HEADER')
    await expect(editor.locator('[data-style="WorkspaceStyle"]')).toHaveCSS('font-size', '24px')
    expect(await readFile(path)).toEqual(Buffer.from(original))
    await panel.getByRole('button', { name: 'Undo this run', exact: true }).click()
    await expect(panel).toContainText('Reverted')
    await expect(docs.locator('.page-hf-header')).not.toContainText(['WORKSPACE_HEADER'])
    await expect(editor.locator('[data-style="WorkspaceStyle"]')).toHaveCount(0)
    await send('Insert a section, change its layout and format the second paragraph', 4)
    expect(requests[3].messages.filter((m) => m.role === 'tool').at(-2)!.content).toContain(
      'section 1',
    )
    await expect(editor).toContainText('Second layout paragraph')
    expect(await readFile(path)).toEqual(Buffer.from(original))

    // A manual layout edit changes no body nodes, but must still invalidate the run's undo.
    await docs.locator('.ribbon-tab', { hasText: /^Layout$/ }).click()
    await docs.getByRole('button', { name: 'Orientation', exact: true }).click()
    await docs.locator('.layout-menu button:not(.active)').click()
    await panel.getByRole('button', { name: 'Undo this run', exact: true }).click()
    await expect(
      panel.getByRole('button', { name: 'Could not restore. Check the document.' }),
    ).toBeDisabled()
    expect(await readFile(path)).toEqual(Buffer.from(original))

    await editor.click()
    await docs.keyboard.press(process.platform === 'darwin' ? 'Meta+s' : 'Control+s')
    await expect
      .poll(async () => {
        const zip = await JSZip.loadAsync(await readFile(path))
        return zip.file('word/styles.xml')!.async('string')
      })
      .toContain('SavedStyle')
    const saved = await JSZip.loadAsync(await readFile(path))
    expect(await saved.file('word/document.xml')!.async('string')).toContain('SavedStyle')
    expect(
      (await saved.file('word/document.xml')!.async('string')).match(/<w:sectPr/g),
    ).toHaveLength(2)
    expect(
      Object.keys(saved.files).filter((name) => /^word\/footer\d+\.xml$/.test(name)),
    ).not.toHaveLength(0)
    expect(requests).toHaveLength(4)
  } finally {
    await closeAndSaveVideo(launched, 'workspace-docs-layout')
    server.closeAllConnections()
    await new Promise<void>((done) => server.close(() => done()))
  }
})
