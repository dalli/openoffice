import { test, expect } from '@playwright/test'
import { _electron as electron } from 'playwright-core'
import { build } from 'vite'
import { mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import type {
  WorkspaceLayout,
  WorkspacePanelLayout,
} from '../apps/shell/src/shared/workspace-layout'
import type { Rect } from '../apps/shell/src/shared/tab-drag-geometry'
import { dividerRect } from '../apps/shell/src/renderer/src/split-view-geometry'
import { clampSplitRatio } from '../apps/shell/src/main/split-layout'

interface Prototype {
  arrange(input: {
    width: number
    height: number
    uiScale: number
    orientation: 'side' | 'stack' | null
    ratio: number
    panel: WorkspacePanelLayout
  }): Promise<{ layout: WorkspaceLayout; views: { bounds: Rect; visible: boolean }[] }>
}

/** Stage 0 native composition test; deliberately independent of subscriptions and editor engines. */
test('shell AI stays clickable outside two native views across size, side, scale and split', async () => {
  const scratch = await mkdtemp(join(tmpdir(), 'ieum-workspace-prototype-'))
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
  await build({
    configFile: false,
    logLevel: 'error',
    esbuild: { jsx: 'automatic' },
    define: { 'process.env.NODE_ENV': '"production"' },
    build: {
      outDir: scratch,
      emptyOutDir: false,
      lib: {
        entry: resolve('e2e/fixtures/workspace-divider.tsx'),
        formats: ['iife'],
        name: 'WorkspaceDividerPrototype',
        fileName: () => 'divider.js',
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
      GENOFFICE_WORKSPACE_PROTOTYPE_DIVIDER: join(scratch, 'divider.js'),
    },
  })
  try {
    const page = await app.firstWindow()
    await expect
      .poll(() =>
        app.evaluate(
          () => !!(globalThis as unknown as { workspacePrototype?: Prototype }).workspacePrototype,
        ),
      )
      .toBe(true)
    for (const size of [
      { width: 1600, height: 900 },
      { width: 1280, height: 720 },
      { width: 1024, height: 768 },
      { width: 1281, height: 721 },
    ]) {
      for (const uiScale of [1, 1.25, 1.5]) {
        for (const side of ['left', 'right'] as const) {
          for (const orientation of [null, 'side', 'stack'] as const) {
            for (const ratio of size.width % 2 ? [0.2, 0.73] : [0.5]) {
              const input = {
                ...size,
                uiScale,
                orientation,
                ratio,
                panel: { side, width: 360, open: true },
              }
              const result = await app.evaluate(
                async (_, input) =>
                  (
                    globalThis as unknown as { workspacePrototype: Prototype }
                  ).workspacePrototype.arrange(input),
                input,
              )
              const { panel, documents } = result.layout
              expect(panel).not.toBeNull()
              for (const view of result.views.filter((view) => view.visible)) {
                expect(view.bounds.x).toBeGreaterThanOrEqual(documents.x)
                // Native rectangles round CSS fractions to integer content pixels.
                expect(view.bounds.x + view.bounds.width).toBeLessThanOrEqual(
                  documents.x + documents.width + 1,
                )
                expect(view.bounds.y).toBeGreaterThanOrEqual(documents.y)
              }
              await expect(page.getByRole('complementary')).toHaveCount(1)
              if (orientation) {
                const gap = dividerRect(
                  { ...size, uiScale, documents },
                  orientation,
                  clampSplitRatio(ratio, documents, orientation),
                )
                const painted = await page.getByRole('separator').boundingBox()
                expect(painted!.x).toBeCloseTo(gap.left)
                expect(painted!.y).toBeCloseTo(gap.top)
                expect(painted!.width).toBeCloseTo(gap.width)
                expect(painted!.height).toBeCloseTo(gap.height)
              } else {
                await expect(page.getByRole('separator')).toHaveCount(0)
              }
              if (result.layout.panelMode === 'open') {
                await page.locator('#fake-response').click()
                await expect(page.getByRole('status')).toContainText(
                  'Local fake result for document B',
                )
                await page.locator('#result').evaluate((node) => {
                  node.textContent = ''
                })
              } else {
                expect(result.layout.panelMode).toBe('rail')
                await page.locator('#rail').click()
                await expect(page.locator('#rail')).toHaveText('AI ✓')
              }
            }
          }
        }
      }
    }
  } finally {
    await app.close()
    await rm(scratch, { recursive: true, force: true })
  }
})
