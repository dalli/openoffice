import { test, expect } from '@playwright/test'
import { createServer } from 'node:http'
import { mkdtemp, readFile, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { build } from 'vite'
import { launchShell, waitForPageWithUrl, closeAndSaveVideo } from './helpers'

declare global {
  interface Window {
    workspaceConnectionModule: typeof import('./fixtures/workspace-connection')
    workspaceConnection: ReturnType<
      typeof import('./fixtures/workspace-connection').installWorkspaceConnection
    >
  }
}

test('workspace connection uses the guarded production stream with a local fake model', async () => {
  test.setTimeout(120_000)
  const calls: string[] = []
  let cancelled = 0
  const server = createServer(async (request, response) => {
    const chunks: Buffer[] = []
    for await (const chunk of request) chunks.push(Buffer.from(chunk))
    const body = Buffer.concat(chunks).toString('utf8')
    calls.push(body)
    response.writeHead(200, { 'Content-Type': 'text/event-stream' })
    if (body.includes('HOLD_LOCAL_REQUEST')) {
      response.write('data: {"choices":[{"delta":{"content":"Waiting"}}]}\n\n')
      response.on('close', () => {
        cancelled++
      })
      return
    }
    response.end(
      'data: {"choices":[{"delta":{"content":"LOCAL_FAKE_ANSWER"}}]}\n\ndata: {"choices":[{"delta":{},"finish_reason":"stop"}],"usage":{"prompt_tokens":17,"completion_tokens":5}}\n\ndata: [DONE]\n\n',
    )
  })
  await new Promise<void>((done) => server.listen(0, '127.0.0.1', done))
  const address = server.address()
  if (!address || typeof address === 'string') throw new Error('No fake endpoint')
  const scratch = await mkdtemp(join(tmpdir(), 'workspace-connection-'))
  const source = join(scratch, 'target.md')
  await writeFile(source, 'LOCAL UNSAVED TARGET\n')
  const userDataDir = await mkdtemp(join(tmpdir(), 'workspace-connection-data-'))
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
  await build({
    configFile: false,
    logLevel: 'error',
    esbuild: { jsx: 'automatic' },
    define: { 'process.env.NODE_ENV': '"production"' },
    build: {
      outDir: scratch,
      emptyOutDir: false,
      lib: {
        entry: resolve('e2e/fixtures/workspace-connection.tsx'),
        formats: ['iife'],
        name: 'workspaceConnectionModule',
        cssFileName: 'connection',
        fileName: () => 'connection.js',
      },
    },
  })
  const launched = await launchShell({
    userDataDir,
    onboardingSeen: true,
    openFile: source,
    recordVideo: false,
    videoDir: 'workspace-connection',
  })
  try {
    const page = await waitForPageWithUrl(launched.app, '/shell/out/renderer/index.html')
    const editor = await waitForPageWithUrl(launched.app, '://markdown/')
    await expect(editor.locator('.ProseMirror').first()).toBeVisible()
    await page.addStyleTag({ content: await readFile(join(scratch, 'connection.css'), 'utf8') })
    await launched.app.evaluate(
      async ({ webContents }, { url, source }) => {
        const contents = webContents.getAllWebContents().find((item) => item.getURL() === url)!
        await contents.executeJavaScript(source)
      },
      {
        url: page.url(),
        source: `${await readFile(join(scratch, 'connection.js'), 'utf8')}\nwindow.workspaceConnectionModule = workspaceConnectionModule; void 0;`,
      },
    )
    await page.evaluate(() => {
      window.workspaceConnection = window.workspaceConnectionModule.installWorkspaceConnection()
    })
    const panel = page.locator('.workspace-ai-surface')
    await expect(panel).toBeVisible()
    await panel.getByRole('button', { name: 'New conversation', exact: true }).click()
    expect(calls).toEqual([])
    await page.evaluate(async () => {
      const panel = window.workspaceConnection.current()!.panel!
      const id = panel.host.conversations.selectedId!
      panel.host.conversations.edit(id, {
        composer: {
          ...panel.host.conversations.state(id).composer,
          text: 'Read this local document',
        },
      })
      await panel.controller.send(id)
    })
    await expect(panel).toContainText('LOCAL_FAKE_ANSWER')
    expect(calls).toHaveLength(1)
    expect(calls[0]).toContain('LOCAL UNSAVED TARGET')
    await expect(panel.locator('.ai-usage-chip')).toContainText('17')
    await page.evaluate(async () => {
      const panel = window.workspaceConnection.current()!.panel!
      const id = panel.host.conversations.selectedId!
      panel.host.conversations.edit(id, {
        composer: { ...panel.host.conversations.state(id).composer, text: 'HOLD_LOCAL_REQUEST' },
      })
      await panel.controller.send(id)
    })
    await expect.poll(() => calls.length).toBe(2)
    await page.evaluate(async () => {
      const panel = window.workspaceConnection.current()!.panel!
      const state = panel.host.conversations.state(panel.host.conversations.selectedId!)
      await panel.host.cancel(state.runs[0]!.run)
    })
    await expect.poll(() => cancelled).toBe(1)
    expect(await readFile(source, 'utf8')).toBe('LOCAL UNSAVED TARGET\n')
    await page.evaluate(() => window.workspaceConnection.dispose())
  } finally {
    await closeAndSaveVideo(launched, 'workspace-connection')
    server.closeAllConnections()
    await new Promise<void>((done) => server.close(() => done()))
  }
})
