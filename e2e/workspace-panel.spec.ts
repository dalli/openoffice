import { expect, test } from '@playwright/test'
import { _electron as electron } from 'playwright-core'
import { build } from 'vite'
import { mkdtemp, readFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import { workspacePanelText as text } from '../apps/shell/src/renderer/src/i18n/strings-workspace-panel'
import type {
  WorkspaceLayout,
  WorkspacePanelLayout,
} from '../apps/shell/src/shared/workspace-layout'

test('actual workspace panel controls stay outside native views at every scale, side, theme and split', async () => {
  const scratch = await mkdtemp(join(tmpdir(), 'ieum-workspace-panel-'))
  const renderer = await build({
    configFile: false,
    logLevel: 'error',
    esbuild: { jsx: 'automatic' },
    define: { 'process.env.NODE_ENV': '"production"' },
    build: {
      outDir: scratch,
      emptyOutDir: false,
      lib: {
        entry: resolve('e2e/fixtures/workspace-panel.tsx'),
        formats: ['iife'],
        name: 'WorkspacePanelFixture',
        cssFileName: 'panel',
        fileName: () => 'panel.js',
      },
    },
  })
  void renderer
  await build({
    configFile: false,
    logLevel: 'error',
    build: {
      outDir: scratch,
      emptyOutDir: false,
      lib: {
        entry: resolve('e2e/fixtures/workspace-layout-model.ts'),
        formats: ['cjs'],
        fileName: () => 'model.cjs',
      },
    },
  })
  const app = await electron.launch({
    args: [
      resolve('e2e/fixtures/workspace-layout-prototype.cjs'),
      `--user-data-dir=${join(scratch, 'profile')}`,
    ],
    env: {
      ...process.env,
      GENOFFICE_WORKSPACE_PROTOTYPE_MODEL: join(scratch, 'model.cjs'),
      GENOFFICE_WORKSPACE_PROTOTYPE_DIVIDER: join(scratch, 'panel.js'),
    },
  })
  try {
    const page = await app.firstWindow()
    await page.addStyleTag({
      content: await readFile(join(scratch, 'panel.css'), 'utf8'),
    })
    await expect
      .poll(() =>
        app.evaluate(() => !!(globalThis as { workspacePrototype?: unknown }).workspacePrototype),
      )
      .toBe(true)
    type Input = {
      width: number
      height: number
      uiScale: number
      orientation: 'side' | 'stack' | null
      ratio: number
      panel: WorkspacePanelLayout
    }
    const arrange = (input: Input) =>
      app.evaluate(
        async (_, input) =>
          (
            globalThis as unknown as {
              workspacePrototype: {
                arrange(input: Input): Promise<{
                  layout: WorkspaceLayout
                  views: Array<{ bounds: { x: number; width: number }; visible: boolean }>
                }>
              }
            }
          ).workspacePrototype.arrange(input),
        input,
      )
    for (const theme of ['light', 'dark']) {
      await page.evaluate(
        (theme) => document.documentElement.setAttribute('data-theme', theme),
        theme,
      )
      for (const uiScale of [1, 1.25, 1.5]) {
        for (const side of ['left', 'right'] as const) {
          for (const orientation of [null, 'side', 'stack'] as const) {
            const result = await arrange({
              width: 1281,
              height: 721,
              uiScale,
              orientation,
              ratio: 0.73,
              panel: { side, width: 360, open: true },
            })
            const panel = page.getByRole('complementary')
            await expect(panel).toHaveCount(1)
            const bounds = (await panel.boundingBox())!
            expect(bounds.x).toBe(result.layout.panel!.x)
            expect(bounds.width).toBe(result.layout.panel!.width)
            for (const view of result.views.filter((item) => item.visible)) {
              expect(view.bounds.x).toBeGreaterThanOrEqual(result.layout.documents.x)
              expect(view.bounds.x + view.bounds.width).toBeLessThanOrEqual(
                result.layout.documents.x + result.layout.documents.width + 1,
              )
            }
            const input = panel.getByRole('textbox')
            await input.fill(`${theme}-${uiScale}-${side}-${orientation}`)
            const inputBounds = (await input.boundingBox())!
            expect(inputBounds.y + inputBounds.height).toBeLessThanOrEqual(bounds.y + bounds.height)
            await expect(input).toBeInViewport()
            await panel.getByRole('combobox', { name: 'Edit target' }).selectOption('B')
            await expect(panel.getByRole('combobox', { name: 'Edit target' })).toHaveValue('B')
          }
        }
      }
    }
    await arrange({
      width: 1024,
      height: 768,
      uiScale: 1.5,
      orientation: 'side',
      ratio: 0.5,
      panel: { side: 'right', width: 360, open: true },
    })
    await expect(page.getByRole('button', { name: 'Open AI Assistant' })).toBeVisible()
    await page.getByRole('button', { name: 'Open AI Assistant' }).click()
    await expect
      .poll(() =>
        page.evaluate(() =>
          (
            window as unknown as { workspacePanelProbe: { requestedOpen(): boolean } }
          ).workspacePanelProbe.requestedOpen(),
        ),
      )
      .toBe(true)
    await arrange({
      width: 1280,
      height: 720,
      uiScale: 1,
      orientation: 'stack',
      ratio: 0.5,
      panel: { side: 'left', width: 360, open: true },
    })
    const firstDraft = await page.getByRole('textbox').inputValue()
    await page.getByRole('button', { name: 'New conversation', exact: true }).click()
    await expect(page.getByRole('textbox')).toHaveValue('')
    await page.getByRole('textbox').fill('second independent draft')
    await page
      .getByRole('combobox', { name: 'Conversation', exact: true })
      .selectOption({ index: 1 })
    await expect(page.getByRole('textbox')).toHaveValue(firstDraft)
    await page.getByRole('combobox', { name: 'Reference document' }).selectOption('A')
    await page.getByRole('combobox', { name: 'Read scope' }).selectOption('selection')
    await page.getByRole('button', { name: 'Add reference' }).click()
    await expect(page.getByRole('button', { name: 'Remove reference · Synthetic A' })).toBeVisible()
    await page.evaluate(() => {
      document.documentElement.lang = 'ar'
      document.documentElement.dir = 'rtl'
    })
    await arrange({
      width: 1280,
      height: 720,
      uiScale: 1.25,
      orientation: 'stack',
      ratio: 0.5,
      panel: { side: 'left', width: 360, open: true },
    })
    await expect(
      page.getByRole('combobox', { name: text('ar', 'target'), exact: true }),
    ).toBeVisible()
    await expect(page.getByRole('textbox')).toHaveValue(firstDraft)
    await page.screenshot({ path: test.info().outputPath('workspace-panel-rtl-dark.png') })
    expect(
      await page.evaluate(() =>
        (
          window as unknown as { workspacePanelProbe: { calls(): number } }
        ).workspacePanelProbe.calls(),
      ),
    ).toBe(0)
  } finally {
    await app.close()
    await rm(scratch, { recursive: true, force: true })
  }
})
