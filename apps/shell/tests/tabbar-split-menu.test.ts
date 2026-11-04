/** @vitest-environment jsdom */
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { TabBar } from '../src/renderer/src/TabBar'
import { LocaleProvider } from '../src/renderer/src/locale'
import type { TabsApi } from '../src/shared/tabs-api'

let root: Root
let container: HTMLDivElement
let finishMenu: () => void
let showSplitMenu: ReturnType<typeof vi.fn>

beforeEach(async () => {
  globalThis.IS_REACT_ACT_ENVIRONMENT = true
  HTMLElement.prototype.scrollIntoView = vi.fn()
  showSplitMenu = vi.fn(
    () =>
      new Promise<void>((resolve) => {
        finishMenu = resolve
      }),
  )
  window.aiOfficeTabs = {
    list: vi.fn(async () => []),
    onChanged: vi.fn(() => () => {}),
    onDockPreview: vi.fn(() => () => {}),
    onSplitHint: vi.fn(() => () => {}),
    notifyChromePressed: vi.fn(),
    showSplitMenu,
  } as unknown as TabsApi
  container = document.createElement('div')
  document.body.append(container)
  root = createRoot(container)
  await act(async () => {
    root.render(
      createElement(LocaleProvider, {
        initial: 'en',
        children: createElement(TabBar, { split: null }),
      }),
    )
  })
})

afterEach(async () => {
  await act(async () => root.unmount())
  container.remove()
  vi.restoreAllMocks()
})

describe('the tab-strip split entry', () => {
  it.each(['ArrowDown', 'ArrowUp'])('opens its native keyboard menu with %s', async (key) => {
    const button = container.querySelector<HTMLButtonElement>('.tab-split-btn')!
    vi.spyOn(button, 'getBoundingClientRect').mockReturnValue({ left: 78, bottom: 40 } as DOMRect)
    expect(button.textContent).toBe('Split view')
    expect(button.getAttribute('aria-haspopup')).toBe('menu')
    expect(button.getAttribute('aria-expanded')).toBe('false')
    await act(async () => {
      button.focus()
      button.dispatchEvent(new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true }))
    })
    expect(showSplitMenu).toHaveBeenCalledWith(78, 40)
    expect(button.getAttribute('aria-expanded')).toBe('true')
    await act(async () => finishMenu())
    expect(button.getAttribute('aria-expanded')).toBe('false')
    expect(document.activeElement).toBe(button)
  })

  it('opens on click and does not open a second menu while one is active', async () => {
    const button = container.querySelector<HTMLButtonElement>('.tab-split-btn')!
    await act(async () => button.click())
    await act(async () => button.click())
    expect(showSplitMenu).toHaveBeenCalledOnce()
    await act(async () => finishMenu())
    await act(async () => button.click())
    expect(showSplitMenu).toHaveBeenCalledTimes(2)
    await act(async () => finishMenu())
  })

  it('lets the button open again after a failed native menu request', async () => {
    const button = container.querySelector<HTMLButtonElement>('.tab-split-btn')!
    showSplitMenu.mockRejectedValueOnce(new Error('menu unavailable'))
    await act(async () => button.click())
    expect(button.getAttribute('aria-expanded')).toBe('false')
    await act(async () => button.click())
    expect(showSplitMenu).toHaveBeenCalledTimes(2)
    await act(async () => finishMenu())
  })
})
