import type { MenuItemConstructorOptions } from 'electron'
import type { SplitDropIntent, SplitOrientation, SplitViewState } from '../shared/split-view-api'
import type { TabSummary } from '../shared/tabs-api'
import type { WorkspaceSplitStringKey } from '../renderer/src/i18n/strings-workspace-split'

/** The existing tab manager owns opening copies, recovery, layout and split persistence. */
export interface SplitMenuHost {
  list(): TabSummary[]
  splitState(): SplitViewState | null
  canDetachTab(id: string): boolean
  splitWith(id: string, intent: SplitDropIntent): void
  beginSplit(
    aTabId: string,
    bTabId: string,
    orientation: SplitOrientation,
    ratio: number,
    focus: 'a' | 'b',
  ): void
  unsplit(): void
}

function eligibleTabs(host: SplitMenuHost): TabSummary[] {
  return host
    .list()
    .filter((tab) => tab.kind !== 'home' && tab.kind !== 'help' && host.canDetachTab(tab.id))
}

/** Re-read eligibility on selection: a native menu may outlive a closing document. */
export function arrangeFromSplitMenu(
  host: SplitMenuHost,
  orientation: SplitOrientation,
  allowCopy = false,
): void {
  const documents = eligibleTabs(host)
  const active = documents.find((tab) => tab.active)
  if (!active) return
  const split = host.splitState()
  if (split && (split.aTabId === active.id || split.bTabId === active.id)) {
    host.splitWith(active.id, { kind: 'flip', orientation })
    return
  }
  const partner = documents.find((tab) => tab.id !== active.id)
  if (partner) {
    host.beginSplit(active.id, partner.id, orientation, 0.5, 'a')
  } else if (allowCopy) {
    host.splitWith(active.id, { kind: 'create', orientation })
  }
}

export function replaceFromSplitMenu(host: SplitMenuHost, side: 'a' | 'b', id: string): void {
  const split = host.splitState()
  const documents = eligibleTabs(host)
  if (!split || !documents.some((tab) => tab.active)) return
  const otherId = side === 'a' ? split.bTabId : split.aTabId
  if (id === otherId || !documents.some((tab) => tab.id === id)) return
  if (!documents.some((tab) => tab.id === otherId)) return
  host.beginSplit(
    side === 'a' ? id : split.aTabId,
    side === 'b' ? id : split.bTabId,
    split.orientation,
    split.ratio,
    side,
  )
}

/** Native menus remain above sibling editor WebContentsViews and provide OS keyboard navigation. */
export function splitMenuTemplate(
  host: SplitMenuHost,
  text: (key: WorkspaceSplitStringKey) => string,
): MenuItemConstructorOptions[] {
  const documents = eligibleTabs(host)
  const enabled = documents.some((tab) => tab.active)
  const split = host.splitState()
  const template: MenuItemConstructorOptions[] = []
  if (!enabled) template.push({ label: text('splitUnavailable'), enabled: false })
  else if (!split && documents.length === 1)
    template.push({ label: text('splitSingleCopy'), enabled: false })
  template.push(
    {
      id: 'split-side',
      label: text('splitSide'),
      type: 'checkbox',
      checked: enabled && split?.orientation === 'side',
      enabled,
      click: () => arrangeFromSplitMenu(host, 'side', !split && documents.length === 1),
    },
    {
      id: 'split-stack',
      label: text('splitStack'),
      type: 'checkbox',
      checked: enabled && split?.orientation === 'stack',
      enabled,
      click: () => arrangeFromSplitMenu(host, 'stack', !split && documents.length === 1),
    },
    { type: 'separator' },
    {
      id: 'split-unsplit',
      label: text('splitUnsplit'),
      enabled: enabled && split !== null,
      click: () => {
        if (eligibleTabs(host).some((tab) => tab.active)) host.unsplit()
      },
    },
  )
  if (enabled && split) {
    template.push({ type: 'separator' })
    for (const side of ['a', 'b'] as const) {
      const otherId = side === 'a' ? split.bTabId : split.aTabId
      const currentId = side === 'a' ? split.aTabId : split.bTabId
      const labelKey =
        split.orientation === 'side'
          ? side === 'a'
            ? 'splitLeftDocument'
            : 'splitRightDocument'
          : side === 'a'
            ? 'splitTopDocument'
            : 'splitBottomDocument'
      template.push({
        id: `split-document-${side}`,
        label: text(labelKey),
        submenu: documents
          .filter((tab) => tab.id !== otherId)
          .map((tab) => ({
            // The strip order distinguishes equally named and untitled documents.
            label: `${host.list().findIndex((candidate) => candidate.id === tab.id)}. ${tab.title}`,
            type: 'checkbox' as const,
            checked: tab.id === currentId,
            click: () => replaceFromSplitMenu(host, side, tab.id),
          })),
      })
    }
  }
  return template
}
