import { test, expect, type Frame } from '@playwright/test'
import { createServer } from 'node:http'
import { copyFile, mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, waitForPageWithUrl, closeAndSaveVideo } from './helpers'

type Studio = {
  plugins: { invoke(id: string, method: string, args: unknown[]): Promise<unknown> }
  automation: {
    registerCommand(command: { id: string; label: string; execute(services: any): void }): void
    execute(id: string): void
  }
}

async function text(frame: Frame): Promise<string> {
  return frame.evaluate(async () => {
    const studio = (window as unknown as { rhwpStudio: Studio }).rhwpStudio
    return (await studio.plugins.invoke('ieum-agent', 'readText', [])) as string
  })
}

async function bytes(frame: Frame): Promise<number[][]> {
  return frame.evaluate(() => {
    const studio = (window as unknown as { rhwpStudio: Studio }).rhwpStudio
    let result: number[][] = []
    const command = `ext:e2e-workspace-bytes-${crypto.randomUUID()}`
    studio.automation.registerCommand({
      id: command,
      label: 'Read fixture bytes',
      execute(services) {
        const doc = services.wasm.borrowDocumentHandle()
        result = [Array.from(doc.exportHwp()), Array.from(doc.exportHwpx())]
      },
    })
    studio.automation.execute(command)
    return result
  })
}

test('the default panel undoes an entire HWP run and refuses undo after another edit', async () => {
  test.setTimeout(120_000)
  let calls = 0
  const server = createServer(async (request, response) => {
    for await (const _chunk of request) {
      /* drain the local request */
    }
    calls++
    const delta =
      calls % 2 === 1
        ? {
            tool_calls: ['FIRST_UNSAVED_MARKER', 'SECOND_UNSAVED_MARKER'].map((marker, index) => ({
              index,
              id: `call-${calls}-${index}`,
              type: 'function',
              function: {
                name: 'insert_paragraphs',
                arguments: JSON.stringify({ where: 'end', paragraphs: [marker] }),
              },
            })),
          }
        : { content: `HWP_RUN_DONE_${calls}` }
    response.writeHead(200, { 'Content-Type': 'text/event-stream' })
    response.end(
      `data: ${JSON.stringify({ choices: [{ delta, finish_reason: calls % 2 ? 'tool_calls' : 'stop' }] })}\n\ndata: [DONE]\n\n`,
    )
  })
  await new Promise<void>((done) => server.listen(0, '127.0.0.1', done))
  const address = server.address()
  if (!address || typeof address === 'string') throw new Error('No local endpoint')
  const scratch = await mkdtemp(join(tmpdir(), 'workspace-hwp-undo-'))
  const path = join(scratch, 'fixture.hwpx')
  await copyFile(join(__dirname, '../apps/hwp/tests/fixtures/memo.hwpx'), path)
  const original = await readFile(path)
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
    videoDir: 'workspace-hwp-undo',
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '/shell/out/renderer/index.html')
    const hwp = await waitForPageWithUrl(launched.app, '://hwp/')
    await hwp.waitForSelector('.hwp-app[data-phase="ready"]')
    const frame = hwp.frames().find((item) => item.url().includes('/studio/'))!
    await frame.evaluate(async () => {
      const studio = (window as unknown as { rhwpStudio: Studio }).rhwpStudio
      await studio.plugins.invoke('ieum-agent', 'insertParagraphs', [
        '',
        null,
        'end',
        ['PRE_RUN_UNSAVED'],
      ])
    })
    const before = await text(frame)
    const beforeBytes = await bytes(frame)
    await page.locator('.tab-split-btn[aria-pressed]').click()
    const panel = page.locator('.workspace-ai-surface')
    await panel.getByRole('button', { name: 'New conversation', exact: true }).click()
    await expect(
      panel.getByRole('combobox', { name: 'Edit target', exact: true }).locator('option:checked'),
    ).toContainText('fixture.hwpx')
    expect(calls).toBe(0)
    await panel.getByRole('textbox').fill('Add two paragraphs to the HWP fixture')
    await panel.getByRole('button', { name: 'Send', exact: true }).click()
    await expect(panel).toContainText('HWP_RUN_DONE_2')
    await expect.poll(() => text(frame)).toContain('SECOND_UNSAVED_MARKER')
    await panel.getByRole('button', { name: 'Undo this run', exact: true }).click()
    await expect(panel).toContainText('Reverted')
    expect(await text(frame)).toBe(before)
    expect(await bytes(frame)).toEqual(beforeBytes)

    // Continue the same conversation and pinned target after undo, then after another edit.
    await panel.getByRole('textbox').fill('Add the paragraphs again')
    await panel.getByRole('button', { name: 'Send', exact: true }).click()
    await expect(panel).toContainText('HWP_RUN_DONE_4')
    await panel.getByRole('textbox').fill('Continue in this conversation once more')
    await panel.getByRole('button', { name: 'Send', exact: true }).click()
    await expect(panel).toContainText('HWP_RUN_DONE_6')
    await frame.evaluate(async () => {
      const studio = (window as unknown as { rhwpStudio: Studio }).rhwpStudio
      await studio.plugins.invoke('ieum-agent', 'insertParagraphs', [
        '',
        null,
        'end',
        ['MANUAL_AFTER_RUN'],
      ])
    })
    const changed = await text(frame)
    await panel.getByRole('button', { name: 'Undo this run', exact: true }).last().click()
    await expect(
      panel.getByRole('button', { name: 'Could not restore. Check the document.' }),
    ).toBeDisabled()
    expect(await text(frame)).toBe(changed)
    expect(await readFile(path)).toEqual(original)
    expect(calls).toBe(6)
  } finally {
    await closeAndSaveVideo(launched, 'workspace-hwp-undo')
    server.closeAllConnections()
    await new Promise<void>((done) => server.close(() => done()))
  }
})
