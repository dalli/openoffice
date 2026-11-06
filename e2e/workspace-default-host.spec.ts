import { test, expect } from '@playwright/test'
import { createServer } from 'node:http'
import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, waitForPageWithUrl, closeAndSaveVideo } from './helpers'

test('the normal shell keeps its workspace host across tabs, collapse, layout changes and cancellation', async () => {
  test.setTimeout(120_000)
  const calls: string[] = []
  let cancelled = 0
  const server = createServer(async (request, response) => {
    const chunks: Buffer[] = []
    for await (const chunk of request) chunks.push(Buffer.from(chunk))
    const body = Buffer.concat(chunks).toString('utf8')
    calls.push(body)
    response.writeHead(200, { 'Content-Type': 'text/event-stream' })
    if (body.includes('HOLD_DEFAULT_HOST')) {
      response.write('data: {"choices":[{"delta":{"content":"Waiting locally"}}]}\n\n')
      response.on('close', () => {
        cancelled++
      })
      return
    }
    response.end(
      'data: {"choices":[{"delta":{"content":"DEFAULT_HOST_ANSWER"}}]}\n\ndata: {"choices":[{"delta":{},"finish_reason":"stop"}],"usage":{"prompt_tokens":17,"completion_tokens":5}}\n\ndata: [DONE]\n\n',
    )
  })
  await new Promise<void>((done) => server.listen(0, '127.0.0.1', done))
  const address = server.address()
  if (!address || typeof address === 'string') throw new Error('No local endpoint')
  const scratch = await mkdtemp(join(tmpdir(), 'workspace-default-'))
  const source = join(scratch, 'first.md')
  const other = join(scratch, 'second.md')
  await writeFile(source, 'FIRST_DOCUMENT_MARKER\n')
  await writeFile(other, 'SECOND_DOCUMENT_MARKER\n')
  const userDataDir = await mkdtemp(join(tmpdir(), 'workspace-default-data-'))
  const config = {
    apiKey: '',
    model: 'local-fixture',
    baseUrl: `http://127.0.0.1:${address.port}/v1`,
  }
  await writeFile(
    join(userDataDir, 'ai-settings.json'),
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
    userDataDir,
    onboardingSeen: true,
    settings: { uiScale: 1.25 },
    openFile: source,
    recordVideo: false,
    videoDir: 'workspace-default-host',
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '/shell/out/renderer/index.html')
    const editor = await waitForPageWithUrl(launched.app, '://markdown/')
    await expect(editor.locator('.ProseMirror').first()).toBeVisible()
    const toggle = page.locator('.tab-split-btn[aria-pressed]')
    await expect(toggle).toBeVisible()
    await expect(page.locator('.workspace-ai-surface')).toHaveCount(0)
    expect(calls).toEqual([])
    await toggle.click()
    const panel = page.locator('.workspace-ai-surface')
    await expect(panel).toBeVisible()
    await panel.getByRole('button', { name: 'New conversation', exact: true }).click()
    const target = panel.getByRole('combobox', { name: 'Edit target', exact: true })
    await expect(target.locator('option:checked')).toContainText('first.md')
    const conversation = panel.getByRole('combobox', { name: 'Conversation', exact: true })
    const first = await conversation.inputValue()
    const input = panel.getByRole('textbox')
    await input.fill('Read the first document')
    await panel.getByRole('button', { name: 'Send', exact: true }).click()
    await expect(panel).toContainText('DEFAULT_HOST_ANSWER')
    expect(calls).toHaveLength(1)
    expect(calls[0]).toContain('FIRST_DOCUMENT_MARKER')
    expect(calls[0]).not.toContain('SECOND_DOCUMENT_MARKER')
    await expect(panel.locator('.ai-usage-chip')).toContainText('17')

    await page.evaluate((path) => window.aiOffice.openPath(path), other)
    await expect
      .poll(
        () => launched.app.windows().filter((item) => item.url().includes('://markdown/')).length,
      )
      .toBe(2)
    const secondEditor = launched.app
      .windows()
      .filter((item) => item.url().includes('://markdown/'))
      .at(-1)!
    await expect(secondEditor.locator('.ProseMirror').first()).toContainText(
      'SECOND_DOCUMENT_MARKER',
    )
    await expect(target.getByRole('option', { name: /second\.md/ })).toBeEnabled()
    await expect(target.locator('option:checked')).toContainText('first.md')
    await expect(panel).toContainText('DEFAULT_HOST_ANSWER')
    await input.fill('FIRST_DRAFT')
    await panel.getByRole('button', { name: 'New conversation', exact: true }).click()
    await expect(target.locator('option:checked')).toContainText('second.md')
    const second = await conversation.inputValue()
    await input.fill('SECOND_DRAFT')
    await conversation.selectOption(first)
    await expect(input).toHaveValue('FIRST_DRAFT')
    await input.fill('HOLD_DEFAULT_HOST')
    await panel.getByRole('button', { name: 'Send', exact: true }).click()
    await expect.poll(() => calls.length).toBe(2)
    await conversation.selectOption(second)
    await expect(input).toHaveValue('SECOND_DRAFT')
    await expect(panel).not.toContainText('Waiting locally')
    await panel.getByRole('button', { name: 'Collapse AI Assistant', exact: true }).click()
    await expect(panel).toHaveCount(0)
    expect(cancelled).toBe(0)
    await toggle.click()
    await expect(input).toHaveValue('SECOND_DRAFT')
    expect(calls).toHaveLength(2)
    await panel
      .getByRole('navigation', { name: 'Running', exact: true })
      .getByRole('button', { name: /^Stop/ })
      .click()
    await expect.poll(() => cancelled).toBe(1)

    // Preference broadcasts move the same host. A constrained window keeps its draft.
    await page.evaluate(() => window.aiOffice.setAiPanelPrefs({ side: 'right', spellcheck: false }))
    await expect(input).toHaveAttribute('spellcheck', 'false')
    await expect.poll(async () => (await panel.boundingBox())!.x).toBeGreaterThan(100)
    await launched.app.evaluate(({ BrowserWindow }) => {
      const win = BrowserWindow.getAllWindows().find((item) =>
        item.webContents.getURL().includes('/shell/'),
      )!
      win.setMinimumSize(0, 0)
      win.setContentSize(560, 720)
    })
    await expect(
      panel.getByRole('button', { name: 'Open AI Assistant', exact: true }),
    ).toBeVisible()
    await launched.app.evaluate(({ BrowserWindow }) => {
      BrowserWindow.getAllWindows()
        .find((win) => win.webContents.getURL().includes('/shell/'))!
        .setContentSize(1280, 800)
    })
    await expect(input).toHaveValue('SECOND_DRAFT')
    await expect.poll(async () => Math.round((await panel.boundingBox())!.width)).toBe(450)
    for (const orientation of ['side', 'stack'] as const) {
      // Choose the real menu action without relying on OS menu automation.
      await launched.app.evaluate(({ Menu }, id) => {
        const popup = Menu.prototype.popup
        Menu.prototype.popup = function (options) {
          Menu.prototype.popup = popup
          const item = this.getMenuItemById(id)
          if (!item?.enabled) throw new Error('Split action unavailable')
          item.click()
          options?.callback?.()
        }
      }, `split-${orientation}`)
      await page.evaluate(() => window.aiOfficeTabs.showSplitMenu(100, 40))
      await expect
        .poll(async () => (await page.evaluate(() => window.aiOfficeTabs.getSplit()))?.orientation)
        .toBe(orientation)
      const bounds = (await panel.boundingBox())!
      const documents = await launched.app.evaluate(({ BrowserWindow, WebContentsView }) => {
        const win = BrowserWindow.getAllWindows().find((item) =>
          item.webContents.getURL().includes('/shell/'),
        )!
        return win.contentView.children
          .filter(
            (view): view is Electron.WebContentsView =>
              view instanceof WebContentsView &&
              view.getVisible() &&
              view.webContents.getURL().includes('://markdown/'),
          )
          .map((view) => view.getBounds())
      })
      expect(documents).toHaveLength(2)
      for (const document of documents) {
        expect(document.x + document.width).toBeLessThanOrEqual(bounds.x)
        expect(document.y).toBeGreaterThanOrEqual(Math.round(bounds.y))
        expect(document.y + document.height).toBeLessThanOrEqual(
          Math.round(bounds.y + bounds.height),
        )
      }
      await expect(page.locator(`.split-divider.${orientation}`)).toBeVisible()
      await expect(input).toHaveValue('SECOND_DRAFT')
      expect(await page.evaluate(() => document.documentElement.scrollHeight)).toBeLessThanOrEqual(
        await page.evaluate(() => window.innerHeight),
      )
    }
    await page.screenshot({ path: test.info().outputPath('workspace-default-host.png') })
    expect(calls).toHaveLength(2)
    expect(await readFile(source, 'utf8')).toBe('FIRST_DOCUMENT_MARKER\n')
    expect(await readFile(other, 'utf8')).toBe('SECOND_DOCUMENT_MARKER\n')
  } finally {
    await closeAndSaveVideo(launched, 'workspace-default-host')
    server.closeAllConnections()
    await new Promise<void>((done) => server.close(() => done()))
  }
})
