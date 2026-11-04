/** @vitest-environment jsdom */
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { isEditingKeyTarget, useModalKeys } from '../src/modal-keys'

function Dialog({ close, nested = false }: { close: () => void; nested?: boolean }) {
  const keys = useModalKeys(close, { restoreFocus: true })
  return createElement(
    'div',
    { ref: keys.ref, onKeyDown: keys.onKeyDown, tabIndex: -1, role: 'dialog' },
    createElement('button', { disabled: true }, 'Disabled'),
    createElement('div', { inert: true }, createElement('button', {}, 'Inactive')),
    createElement('input', { 'aria-label': 'First' }),
    createElement('a', { href: '#last' }, 'Last'),
    nested ? createElement(Dialog, { close }) : null,
  )
}

describe('shared modal focus', () => {
  let host: HTMLDivElement
  let root: ReturnType<typeof createRoot>
  let opener: HTMLButtonElement

  beforeEach(() => {
    vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true)
    opener = document.createElement('button')
    host = document.createElement('div')
    document.body.append(opener, host)
    opener.focus()
    root = createRoot(host)
  })

  afterEach(() => {
    act(() => root.unmount())
    host.remove()
    opener.remove()
    vi.unstubAllGlobals()
  })

  it('enters the first available control, blocks background focus, and restores the opener', () => {
    act(() => root.render(createElement(Dialog, { close: vi.fn() })))
    const first = host.querySelector('input')!
    expect(document.activeElement).toBe(first)
    act(() => opener.focus())
    expect(document.activeElement).toBe(first)
    act(() => root.render(null))
    expect(document.activeElement).toBe(opener)
  })

  it('cycles inputs and links while skipping disabled and inert controls', () => {
    act(() => root.render(createElement(Dialog, { close: vi.fn() })))
    const first = host.querySelector('input')!
    const last = host.querySelector('a')!
    act(() =>
      first.dispatchEvent(
        new KeyboardEvent('keydown', {
          key: 'Tab',
          shiftKey: true,
          bubbles: true,
          cancelable: true,
        }),
      ),
    )
    expect(document.activeElement).toBe(last)
    act(() =>
      last.dispatchEvent(
        new KeyboardEvent('keydown', { key: 'Tab', bubbles: true, cancelable: true }),
      ),
    )
    expect(document.activeElement).toBe(first)
  })

  it('lets the innermost modal own focus and Escape even when both mount together', () => {
    const close = vi.fn()
    act(() => root.render(createElement(Dialog, { close, nested: true })))
    const dialogs = host.querySelectorAll('[role="dialog"]')
    expect(dialogs[1]?.contains(document.activeElement)).toBe(true)
    act(() =>
      (document.activeElement as HTMLElement).dispatchEvent(
        new KeyboardEvent('keydown', { key: 'Escape', bubbles: true, cancelable: true }),
      ),
    )
    expect(close).toHaveBeenCalledTimes(1)
  })

  it('recognizes form fields and inherited editable content', () => {
    const controls = document.createElement('div')
    controls.innerHTML =
      '<input><textarea></textarea><select></select><div contenteditable="true"><span>x</span></div><button>Next</button>'
    for (const selector of ['input', 'textarea', 'select', 'span']) {
      expect(isEditingKeyTarget(controls.querySelector(selector))).toBe(true)
    }
    expect(isEditingKeyTarget(controls.querySelector('button'))).toBe(false)
    expect(isEditingKeyTarget(window)).toBe(false)
  })
})
