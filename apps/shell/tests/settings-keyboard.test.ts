/** @vitest-environment jsdom */
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import type { HomeApi } from '../src/shared/home-api'
import { LocaleProvider } from '../src/renderer/src/locale'
import { SettingsModal } from '../src/renderer/src/SettingsModal'

describe('Settings keyboard focus', () => {
  let host: HTMLDivElement
  let root: ReturnType<typeof createRoot>
  let opener: HTMLButtonElement

  beforeEach(() => {
    vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true)
    window.aiOffice = {
      getTheme: async () => 'system',
      getDefaultSaveDir: async () => '',
      getUpdateChannel: async () => 'stable',
      getAppVersion: async () => '1.0.0',
    } as unknown as HomeApi
    opener = document.createElement('button')
    host = document.createElement('div')
    document.body.append(opener, host)
    opener.focus()
    root = createRoot(host)
  })

  afterEach(() => {
    act(() => root.unmount())
    opener.remove()
    host.remove()
    vi.unstubAllGlobals()
  })

  const key = (target: HTMLElement, value: string, shiftKey = false) => {
    act(() =>
      target.dispatchEvent(
        new KeyboardEvent('keydown', { key: value, shiftKey, bubbles: true, cancelable: true }),
      ),
    )
  }

  it('enters settings, wraps both directions and restores focus after Escape', async () => {
    await act(async () =>
      root.render(
        createElement(
          LocaleProvider,
          { initial: 'en' },
          createElement(SettingsModal, { onClose: () => root.render(null) }),
        ),
      ),
    )
    const close = host.querySelector<HTMLButtonElement>('.set-close')!
    expect(document.activeElement).toBe(close)
    key(close, 'Tab', true)
    const last = document.activeElement as HTMLElement
    expect(last).not.toBe(close)
    expect(host.querySelector('.set-dialog')?.contains(last)).toBe(true)
    key(last, 'Tab')
    expect(document.activeElement).toBe(close)
    act(() => opener.focus())
    expect(document.activeElement).toBe(close)
    key(close, 'Escape')
    expect(host.querySelector('.set-dialog')).toBeNull()
    expect(document.activeElement).toBe(opener)
  })

  it('lets an open settings dropdown consume Escape before the modal', async () => {
    const close = vi.fn()
    await act(async () =>
      root.render(
        createElement(
          LocaleProvider,
          { initial: 'en' },
          createElement(SettingsModal, { onClose: close, target: { section: 'general' } }),
        ),
      ),
    )
    const trigger = host.querySelector<HTMLButtonElement>('.gs-dd-btn')!
    act(() => {
      trigger.focus()
      trigger.click()
    })
    expect(host.querySelector('.gs-dd-pop')).not.toBeNull()
    key(trigger, 'Escape')
    expect(host.querySelector('.gs-dd-pop')).toBeNull()
    expect(close).not.toHaveBeenCalled()
    key(trigger, 'Escape')
    expect(close).toHaveBeenCalledTimes(1)
  })
})
