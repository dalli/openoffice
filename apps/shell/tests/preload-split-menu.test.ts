import { describe, expect, it, vi } from 'vitest'
import { TABS_CHANNELS, type TabsApi } from '../src/shared/tabs-api'

const mocks = vi.hoisted(() => ({
  exposed: new Map<string, unknown>(),
  invoke: vi.fn(),
}))

vi.mock('electron', () => ({
  contextBridge: {
    exposeInMainWorld: (name: string, api: unknown) => mocks.exposed.set(name, api),
  },
  ipcRenderer: {
    invoke: mocks.invoke,
    send: vi.fn(),
    on: vi.fn(),
    removeListener: vi.fn(),
  },
  webUtils: { getPathForFile: () => '' },
}))

import '../src/preload/index'

describe('split menu preload bridge', () => {
  it('uses the dedicated channel and waits until the native menu has closed', async () => {
    const api = mocks.exposed.get('aiOfficeTabs') as TabsApi
    let closeMenu!: () => void
    mocks.invoke.mockImplementationOnce(
      () =>
        new Promise<void>((resolve) => {
          closeMenu = resolve
        }),
    )
    let finished = false
    const pending = api.showSplitMenu(120, 40).then(() => {
      finished = true
    })
    await Promise.resolve()
    expect(mocks.invoke).toHaveBeenCalledWith(TABS_CHANNELS.showSplitMenu, 120, 40)
    expect(finished).toBe(false)
    closeMenu()
    await pending
    expect(finished).toBe(true)
  })
})
