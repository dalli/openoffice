import { describe, expect, it, vi } from 'vitest'
import type { MenuItemConstructorOptions } from 'electron'
import { LANGS } from '@genoffice/i18n'
import type { TabSummary } from '../src/shared/tabs-api'
import type { SplitViewState } from '../src/shared/split-view-api'
import {
  arrangeFromSplitMenu,
  replaceFromSplitMenu,
  splitMenuTemplate,
  type SplitMenuHost,
} from '../src/main/split-menu'
import {
  workspaceSplitStrings,
  workspaceSplitText,
} from '../src/renderer/src/i18n/strings-workspace-split'

const tab = (id: string, active = false, kind: TabSummary['kind'] = 'docs'): TabSummary => ({
  id,
  kind,
  title: id,
  closable: kind !== 'home',
  active,
})
const home = tab('home', false, 'home')
const split: SplitViewState = {
  aTabId: 'a',
  bTabId: 'b',
  orientation: 'side',
  ratio: 0.35,
  focus: 'b',
}
const text = (key: keyof typeof workspaceSplitStrings.zh) => workspaceSplitText('en', key)

function host(tabs: TabSummary[], state: SplitViewState | null = null) {
  return {
    list: vi.fn(() => tabs),
    splitState: vi.fn(() => state),
    canDetachTab: vi.fn((id: string) => id !== 'home' && id !== 'present'),
    splitWith: vi.fn(),
    beginSplit: vi.fn(),
    unsplit: vi.fn(),
  } satisfies SplitMenuHost
}
const click = (item: MenuItemConstructorOptions): void => (item.click as unknown as () => void)()
const item = (menu: MenuItemConstructorOptions[], id: string): MenuItemConstructorOptions =>
  menu.find((entry) => entry.id === id)!

describe('window split menu', () => {
  it('exposes both directions and unsplit in one menu', () => {
    const menu = splitMenuTemplate(host([home, tab('a', true), tab('b')]), text)
    expect(item(menu, 'split-side')).toMatchObject({ label: 'Side by side', enabled: true })
    expect(item(menu, 'split-stack')).toMatchObject({ label: 'Top and bottom', enabled: true })
    expect(item(menu, 'split-unsplit')).toMatchObject({ label: 'Unsplit', enabled: false })
  })

  it('uses two open documents even after Home or help was most recently selected', () => {
    const manager = host([home, tab('a'), tab('help', false, 'help'), tab('b', true)])
    click(item(splitMenuTemplate(manager, text), 'split-side'))
    expect(manager.beginSplit).toHaveBeenCalledWith('b', 'a', 'side', 0.5, 'a')
    expect(manager.splitWith).not.toHaveBeenCalled()
  })

  it('explains a single-document copy before delegating to the existing copy guard', () => {
    const manager = host([home, tab('a', true)])
    const menu = splitMenuTemplate(manager, text)
    expect(menu[0]).toEqual({ label: text('splitSingleCopy'), enabled: false })
    click(item(menu, 'split-stack'))
    expect(manager.splitWith).toHaveBeenCalledWith('a', { kind: 'create', orientation: 'stack' })
    expect(manager.beginSplit).not.toHaveBeenCalled()
  })

  it('does not silently open a copy if the other document closes while the menu is open', () => {
    const manager = host([home, tab('a', true), tab('b')])
    const menu = splitMenuTemplate(manager, text)
    manager.list.mockReturnValue([home, tab('a', true)])
    click(item(menu, 'split-side'))
    expect(manager.splitWith).not.toHaveBeenCalled()
    expect(manager.beginSplit).not.toHaveBeenCalled()
  })

  it('preserves both documents and the ratio when changing split direction', () => {
    const manager = host([home, tab('a'), tab('b', true)], split)
    const menu = splitMenuTemplate(manager, text)
    expect(item(menu, 'split-side').checked).toBe(true)
    expect(item(menu, 'split-stack').checked).toBe(false)
    click(item(menu, 'split-stack'))
    expect(manager.splitWith).toHaveBeenCalledWith('b', { kind: 'flip', orientation: 'stack' })
    expect(manager.beginSplit).not.toHaveBeenCalled()
  })

  it.each(['home', 'help', 'present'])('explains why %s cannot be split', (id) => {
    const active = tab(id, true, id === 'present' ? 'html' : id === 'help' ? 'help' : 'home')
    const manager = host([active, tab('a'), tab('b')], split)
    const menu = splitMenuTemplate(manager, text)
    expect(menu[0]).toEqual({ label: text('splitUnavailable'), enabled: false })
    for (const id of ['split-side', 'split-stack', 'split-unsplit']) {
      expect(item(menu, id).enabled).toBe(false)
      click(item(menu, id))
    }
    expect(manager.beginSplit).not.toHaveBeenCalled()
    expect(manager.splitWith).not.toHaveBeenCalled()
    expect(manager.unsplit).not.toHaveBeenCalled()
  })

  it('offers per-pane replacement and distinguishes equally named documents by strip order', () => {
    const tabs = [home, tab('a'), tab('b', true), { ...tab('c'), title: 'a' }]
    const manager = host(tabs, split)
    const menu = splitMenuTemplate(manager, text)
    const submenu = item(menu, 'split-document-a').submenu as MenuItemConstructorOptions[]
    expect(item(menu, 'split-document-a').label).toBe('Left document')
    expect(item(menu, 'split-document-b').label).toBe('Right document')
    expect(submenu.map((entry) => entry.label)).toEqual(['1. a', '3. a'])
    click(submenu[1])
    expect(manager.beginSplit).toHaveBeenCalledWith('c', 'b', 'side', 0.35, 'a')
    expect(manager.unsplit).not.toHaveBeenCalled()
  })

  it('names stacked panes top and bottom and keeps their order when replacing the lower one', () => {
    const manager = host([home, tab('a'), tab('b', true), tab('c')], {
      ...split,
      orientation: 'stack',
    })
    const menu = splitMenuTemplate(manager, text)
    expect(item(menu, 'split-document-a').label).toBe('Top document')
    expect(item(menu, 'split-document-b').label).toBe('Bottom document')
    replaceFromSplitMenu(manager, 'b', 'c')
    expect(manager.beginSplit).toHaveBeenCalledWith('a', 'c', 'stack', 0.35, 'b')
  })

  it('rejects stale replacements, missing opposite documents and reuse of the same tab', () => {
    const manager = host([home, tab('a'), tab('b', true), tab('c')], split)
    replaceFromSplitMenu(manager, 'a', 'b')
    replaceFromSplitMenu(manager, 'a', 'missing')
    manager.list.mockReturnValue([home, tab('a', true), tab('c')])
    replaceFromSplitMenu(manager, 'a', 'c')
    expect(manager.beginSplit).not.toHaveBeenCalled()
  })

  it('unsplits through the existing manager without closing or saving documents', () => {
    const manager = host([home, tab('a'), tab('b', true)], split)
    click(item(splitMenuTemplate(manager, text), 'split-unsplit'))
    expect(manager.unsplit).toHaveBeenCalledOnce()
    expect(manager.beginSplit).not.toHaveBeenCalled()
    expect(manager.splitWith).not.toHaveBeenCalled()
  })

  it('does not arrange a document that has left the eligible list', () => {
    const manager = host([home, tab('a', true), tab('b')])
    manager.canDetachTab.mockReturnValue(false)
    arrangeFromSplitMenu(manager, 'side', true)
    expect(manager.beginSplit).not.toHaveBeenCalled()
    expect(manager.splitWith).not.toHaveBeenCalled()
  })

  it.each(LANGS)('has every split string translated into %s', (lang) => {
    const strings = workspaceSplitStrings[lang]
    expect(Object.keys(strings).sort()).toEqual(Object.keys(workspaceSplitStrings.zh).sort())
    for (const key of Object.keys(workspaceSplitStrings.zh) as (keyof typeof strings)[]) {
      expect(strings[key].length).toBeGreaterThan(0)
      if (lang !== 'en') expect(strings[key], key).not.toBe(workspaceSplitStrings.en[key])
    }
  })
})
