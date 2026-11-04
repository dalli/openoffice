import { SPLIT_DIVIDER_PX, SPLIT_MIN_PANE_PX, type SplitOrientation } from './split-view-api'
import { tabStripHeight, type Rect } from './tab-drag-geometry'

/** Panel preferences are logical pixels; returned rectangles are window content pixels. */
export interface WorkspacePanelLayout {
  side: 'left' | 'right'
  width: number
  open: boolean
}

export interface WorkspaceViewport {
  width: number
  height: number
  uiScale?: number
}

export interface WorkspaceLayout {
  documents: Rect
  panel: Rect | null
  panelMode: 'hidden' | 'open' | 'rail'
  /** A temporary space constraint never rewrites the saved panel preference. */
  constrained: boolean
}

export const WORKSPACE_PANEL_DEFAULT_PX = 360
export const WORKSPACE_PANEL_MIN_PX = 300
export const WORKSPACE_PANEL_MAX_PX = 640
export const WORKSPACE_PANEL_RAIL_PX = 40

function nonnegative(value: number): number {
  return Number.isFinite(value) ? Math.max(0, value) : 0
}

export function workspaceUiScale(value = 1): number {
  return Number.isFinite(value) ? Math.min(1.5, Math.max(1, value)) : 1
}

/**
 * Reserve the shell panel before laying out either document. Nothing here
 * opens a panel, reads a document, persists settings or starts a model call.
 * A null panel preserves today's editor-owned AI layout during migration.
 */
export function workspaceLayout(
  viewport: WorkspaceViewport,
  panel: Readonly<WorkspacePanelLayout> | null = null,
  orientation: SplitOrientation | null = null,
): WorkspaceLayout {
  const width = nonnegative(viewport.width)
  const height = nonnegative(viewport.height)
  const scale = workspaceUiScale(viewport.uiScale)
  const strip = tabStripHeight(scale)
  const documents = { x: 0, y: strip, width, height: Math.max(0, height - strip) }
  if (!panel) return { documents, panel: null, panelMode: 'hidden', constrained: false }

  const minDocuments =
    orientation === 'side' ? 2 * SPLIT_MIN_PANE_PX + SPLIT_DIVIDER_PX : SPLIT_MIN_PANE_PX
  const requested = Number.isFinite(panel.width) ? panel.width : WORKSPACE_PANEL_DEFAULT_PX
  const openWidth = Math.round(
    Math.min(WORKSPACE_PANEL_MAX_PX, Math.max(WORKSPACE_PANEL_MIN_PX, requested)) * scale,
  )
  const railWidth = Math.round(WORKSPACE_PANEL_RAIL_PX * scale)
  const canOpen = width >= minDocuments + openWidth
  const canRail = width >= minDocuments + railWidth
  const panelMode = panel.open && canOpen ? 'open' : canRail ? 'rail' : 'hidden'
  const panelWidth = panelMode === 'open' ? openWidth : panelMode === 'rail' ? railWidth : 0
  documents.width = width - panelWidth
  if (panel.side === 'left') documents.x = panelWidth
  return {
    documents,
    panel:
      panelWidth === 0
        ? null
        : {
            x: panel.side === 'left' ? 0 : documents.width,
            y: strip,
            width: panelWidth,
            height: documents.height,
          },
    panelMode,
    constrained: panel.open ? !canOpen : !canRail,
  }
}
