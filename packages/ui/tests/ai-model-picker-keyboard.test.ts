/** @vitest-environment jsdom */
import { act, createElement } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { defaultAiSettings, type AiSettings } from '@genoffice/ai-provider'
import { AiModelPicker, type AiModelPickerBridge } from '../src/AiModelPicker'

describe('AiModelPicker keyboard navigation', () => {
  let host: HTMLDivElement
  let root: ReturnType<typeof createRoot>
  let written: AiSettings[]
  let stored: AiSettings

  beforeEach(() => {
    vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true)
    host = document.createElement('div')
    document.body.append(host)
    root = createRoot(host)
    written = []
    stored = defaultAiSettings()
    stored.provider = 'claude-code'
    stored.providers['claude-code'] = { apiKey: '', model: 'opus' }
  })

  afterEach(() => {
    act(() => root.unmount())
    host.remove()
    vi.unstubAllGlobals()
  })

  async function render() {
    const bridge: AiModelPickerBridge = {
      getSettings: async () => stored,
      setSettings: async (value) => {
        stored = value
        written.push(value)
      },
      openModelSettings: vi.fn(),
    }
    await act(async () => {
      root.render(
        createElement(AiModelPicker, {
          bridge,
          lang: 'en',
          subscriptionStatuses: {
            'claude-code': {
              provider: 'claude-code',
              state: 'ready',
              meteredAccepted: false,
              models: [
                { id: 'haiku', label: 'Haiku' },
                { id: 'opus', label: 'Opus with a long accessible model name', efforts: ['high'] },
              ],
            },
          },
        }),
      )
    })
    return host.querySelector<HTMLButtonElement>('.ai-model-chip')!
  }

  const key = async (target: HTMLElement, value: string, init: KeyboardEventInit = {}) => {
    await act(async () => {
      target.dispatchEvent(
        new KeyboardEvent('keydown', { key: value, bubbles: true, cancelable: true, ...init }),
      )
    })
  }
  const rows = () => [...host.querySelectorAll<HTMLButtonElement>('[role="option"]')]

  it('focuses the selected model and navigates without changing stored settings', async () => {
    const chip = await render()
    act(() => chip.click())
    const selected = host.querySelector<HTMLButtonElement>('[aria-selected="true"]')!
    expect(document.activeElement).toBe(selected)
    expect(selected.title).toContain('Opus with a long accessible model name')
    expect(selected.getAttribute('aria-label')).toContain('Claude')
    expect(document.getElementById(chip.getAttribute('aria-controls')!)?.getAttribute('role')).toBe(
      'dialog',
    )

    await key(selected, 'Home')
    expect(document.activeElement).toBe(rows()[0])
    await key(document.activeElement as HTMLElement, 'ArrowDown')
    expect(document.activeElement).toBe(rows()[1])
    await key(document.activeElement as HTMLElement, 'End')
    expect(document.activeElement).toBe(rows().at(-1))
    await key(document.activeElement as HTMLElement, 'ArrowUp')
    expect(document.activeElement).toBe(rows().at(-2))
    expect(written).toEqual([])
  })

  it('opens from an arrow, selects with Enter, and returns focus to the chip', async () => {
    const chip = await render()
    act(() => chip.focus())
    await key(chip, 'ArrowDown')
    await key(document.activeElement as HTMLElement, 'Home')
    await key(document.activeElement as HTMLElement, 'ArrowDown')
    const target = document.activeElement as HTMLButtonElement
    expect(target.textContent).toContain('Default')
    await key(target, 'Enter')
    expect(written).toHaveLength(1)
    expect(written[0]?.providers['claude-code']?.model).toBe('')
    expect(host.querySelector('.ai-model-pop')).toBeNull()
    expect(document.activeElement).toBe(chip)
  })

  it('Escape restores focus without saving or reaching a hosting modal', async () => {
    const chip = await render()
    const escaped = vi.fn()
    window.addEventListener('keydown', escaped)
    try {
      act(() => chip.click())
      await key(document.activeElement as HTMLElement, 'Escape')
      expect(document.activeElement).toBe(chip)
      expect(chip.getAttribute('aria-expanded')).toBe('false')
      expect(written).toEqual([])
      expect(escaped).not.toHaveBeenCalled()
    } finally {
      window.removeEventListener('keydown', escaped)
    }
  })

  it('keeps effort controls reachable and closes when focus leaves the picker', async () => {
    const chip = await render()
    act(() => chip.click())
    const effort = host.querySelector<HTMLButtonElement>('.ai-effort-option')!
    act(() => effort.focus())
    await key(effort, 'ArrowDown')
    expect(document.activeElement).toBe(effort)
    const outside = document.createElement('button')
    document.body.append(outside)
    try {
      act(() => outside.focus())
      expect(host.querySelector('.ai-model-pop')).toBeNull()
      expect(document.activeElement).toBe(outside)
      expect(written).toEqual([])
    } finally {
      outside.remove()
    }
  })
})
