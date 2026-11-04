import { describe, expect, it } from 'vitest'
import { paneRects, splitDropZoneInRect } from '../src/main/split-layout'
import { dividerRect, ratioForDividerOffset } from '../src/renderer/src/split-view-geometry'
import { tabStripHeight } from '../src/shared/tab-drag-geometry'
import { workspaceLayout, type WorkspacePanelLayout } from '../src/shared/workspace-layout'

describe('workspace document and AI rectangles', () => {
  for (const size of [
    { width: 1600, height: 900 },
    { width: 1280, height: 720 },
    { width: 1024, height: 768 },
  ]) {
    for (const uiScale of [1, 1.25, 1.5]) {
      for (const side of ['left', 'right'] as const) {
        for (const orientation of ['side', 'stack'] as const) {
          it(`${size.width} at ${uiScale}: ${side} AI and ${orientation} split never overlap`, () => {
            const viewport = { ...size, uiScale }
            const layout = workspaceLayout(viewport, { side, width: 360, open: true }, orientation)
            const { documents, panel } = layout
            expect(panel).not.toBeNull()
            expect(panel!.y).toBe(tabStripHeight(uiScale))
            expect(panel!.height).toBe(size.height - tabStripHeight(uiScale))
            expect(documents.width + panel!.width).toBe(size.width)
            expect(side === 'left' ? documents.x : documents.x + documents.width).toBe(
              side === 'left' ? panel!.width : panel!.x,
            )
            const panes = paneRects(documents, orientation, 0.5)
            const gap = dividerRect({ ...viewport, documents }, orientation, 0.5)
            expect(gap.left).toBeGreaterThanOrEqual(documents.x)
            expect(gap.top).toBeGreaterThanOrEqual(documents.y)
            expect(gap.left + gap.width).toBeLessThanOrEqual(documents.x + documents.width)
            expect(gap.top + gap.height).toBeLessThanOrEqual(documents.y + documents.height)
            expect(
              ratioForDividerOffset(
                orientation === 'side' ? gap.left : gap.top,
                { ...viewport, documents },
                orientation,
              ),
            ).toBeCloseTo(0.5)
            if (orientation === 'side') {
              expect(gap.left).toBe(panes.a.x + panes.a.width)
              expect(gap.left + gap.width).toBe(panes.b.x)
            } else {
              expect(gap.top).toBe(panes.a.y + panes.a.height)
              expect(gap.top + gap.height).toBe(panes.b.y)
            }
            expect(splitDropZoneInRect(panel!.x + 1, panel!.y + 1, documents)).toBeNull()
            expect(
              splitDropZoneInRect(documents.x + documents.width - 1, documents.y + 20, documents),
            ).toBe('side')
          })
        }
      }
    }
  }

  it('keeps the legacy full document bounds until a shell host attaches', () => {
    expect(workspaceLayout({ width: 1280, height: 720, uiScale: 1.5 })).toEqual({
      documents: { x: 0, y: 60, width: 1280, height: 660 },
      panel: null,
      panelMode: 'hidden',
      constrained: false,
    })
  })

  it('temporarily uses a rail and restores the requested width without mutating preferences', () => {
    const panel: WorkspacePanelLayout = { side: 'right', width: 600, open: true }
    expect(workspaceLayout({ width: 1024, height: 768 }, panel, 'side')).toMatchObject({
      panelMode: 'rail',
      panel: { width: 40 },
      constrained: true,
    })
    expect(workspaceLayout({ width: 1600, height: 900 }, panel, 'side')).toMatchObject({
      panelMode: 'open',
      panel: { width: 600 },
      constrained: false,
    })
    expect(panel).toEqual({ side: 'right', width: 600, open: true })
  })

  it('gives the documents all available space when even a rail cannot fit', () => {
    expect(
      workspaceLayout(
        { width: 500, height: 400 },
        { side: 'left', width: 360, open: true },
        'side',
      ),
    ).toMatchObject({ documents: { x: 0, width: 500 }, panel: null, constrained: true })
  })

  it('reserves only the rail when collapsed and tolerates corrupt dimensions', () => {
    expect(
      workspaceLayout({ width: 1280, height: 720 }, { side: 'right', width: 600, open: false }),
    ).toMatchObject({ panelMode: 'rail', panel: { width: 40 }, constrained: false })
    expect(
      workspaceLayout(
        { width: 1280, height: 720, uiScale: Number.NaN },
        { side: 'right', width: Number.NaN, open: true },
      ),
    ).toMatchObject({ panel: { width: 360 } })
    expect(workspaceLayout({ width: Number.NaN, height: -5 })).toMatchObject({
      documents: { width: 0, height: 0 },
    })
  })

  it('rejects drops outside the document rather than treating the AI or another window as its edge', () => {
    const { documents } = workspaceLayout(
      { width: 1280, height: 720, uiScale: 1.5 },
      { side: 'right', width: 360, open: true },
    )
    expect(splitDropZoneInRect(1279, 300, documents)).toBeNull()
    expect(splitDropZoneInRect(100, 55, documents)).toBeNull()
    expect(splitDropZoneInRect(-1, 700, documents)).toBeNull()
    expect(splitDropZoneInRect(100, 720, documents)).toBeNull()
  })
})
