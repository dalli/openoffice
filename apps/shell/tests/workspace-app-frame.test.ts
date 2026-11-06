/** @vitest-environment jsdom */
import { act, createElement, StrictMode } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { applyAiPanelPrefs, DEFAULT_AI_PANEL_PREFS, type AiPanelPrefs } from '@genoffice/ui'
import { WorkspaceAppFrame } from '../src/renderer/src/workspace-ai/WorkspaceAppFrame'
import { AppFrameView } from '../src/renderer/src/AppFrame'
import { useWorkspaceConnection } from '../src/renderer/src/workspace-ai/use-workspace-connection'
import type { WorkspacePanelEvent } from '../src/shared/workspace-panel-api'
import { workspaceLayout } from '../src/shared/workspace-layout'

vi.mock('../src/renderer/src/AppFrame', () => ({ AppFrameView: vi.fn(() => null) }))
vi.mock('../src/renderer/src/subscription-client', () => ({
  useSubscriptionConnection: () => ({ statuses: {} }),
}))
vi.mock('../src/renderer/src/workspace-ai/use-workspace-connection', () => ({
  useWorkspaceConnection: vi.fn(),
}))
;(
  globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
).IS_REACT_ACT_ENVIRONMENT = true
let root: Root | null = null
afterEach(async () => {
  await act(async () => root?.unmount())
  root = null
  localStorage.clear()
  applyAiPanelPrefs(DEFAULT_AI_PANEL_PREFS)
  vi.clearAllMocks()
})

async function fixture(seen = true, strict = false) {
  let pushPrefs!: (prefs: AiPanelPrefs) => void
  let resolvePrefs!: (prefs: AiPanelPrefs) => void
  const prefs = new Promise<AiPanelPrefs>((resolve) => {
    resolvePrefs = resolve
  })
  const listeners = new Set<(event: WorkspacePanelEvent) => void>()
  const client = {
    entriesReady: vi.fn(async (_ready: boolean) => {}),
    layout: vi.fn(async (panel) => workspaceLayout({ width: 1280, height: 720 }, panel)),
    subscribe: (listener: (event: WorkspacePanelEvent) => void) => {
      listeners.add(listener)
      return () => {
        listeners.delete(listener)
      }
    },
  }
  const panel = {
    client,
    controller: { enter: vi.fn(async () => {}) },
    dispose: vi.fn(async () => {}),
  }
  const reconnect = vi.fn()
  vi.mocked(useWorkspaceConnection).mockImplementation(
    (options) =>
      ({
        panel: options.enabled ? panel : null,
        failed: false,
        reconnect,
        composerControls: null,
        notices: null,
      }) as unknown as ReturnType<typeof useWorkspaceConnection>,
  )
  Object.assign(window, {
    aiOffice: {
      getAiPanelPrefs: vi.fn(() => prefs),
      onAiPanelPrefsChanged: (listener: typeof pushPrefs) => {
        pushPrefs = listener
        return () => {}
      },
    },
  })
  const props = () => vi.mocked(AppFrameView).mock.calls.at(-1)![0]
  const container = document.createElement('div')
  root = createRoot(container)
  await act(async () =>
    root!.render(
      strict
        ? createElement(
            StrictMode,
            null,
            createElement(WorkspaceAppFrame, { initialOnboardingSeen: seen }),
          )
        : createElement(WorkspaceAppFrame, { initialOnboardingSeen: seen }),
    ),
  )
  return {
    props,
    panel,
    client,
    listeners,
    reconnect,
    async prefs(next: AiPanelPrefs) {
      await act(async () => pushPrefs(next))
    },
    async readPrefs(next = DEFAULT_AI_PANEL_PREFS) {
      await act(async () => resolvePrefs(next))
    },
  }
}

describe('normal workspace shell host', () => {
  it('opens an editor entry without retargeting on collapse and unregisters readiness on teardown', async () => {
    const f = await fixture()
    await f.readPrefs()
    expect(f.client.entriesReady).toHaveBeenLastCalledWith(true)
    const capture = {
      id: 'captured-source',
      binding: {},
    } as import('../src/main/workspace-ai/execution-controller').WorkspaceCapture
    const entry: WorkspacePanelEvent = {
      sessionId: 'session',
      kind: 'entry',
      capture,
      entry: { kind: 'open' },
    }
    await act(async () => {
      for (const listener of f.listeners) listener(entry)
    })
    expect(f.props().workspaceToggle?.open).toBe(true)
    expect(f.panel.controller.enter).toHaveBeenCalledWith(capture, { kind: 'open' }, undefined)
    f.panel.controller.enter.mockClear()
    await act(async () => {
      for (const listener of f.listeners) listener({ ...entry, entry: { kind: 'toggle' } })
    })
    expect(f.props().workspaceToggle?.open).toBe(false)
    expect(f.panel.controller.enter).not.toHaveBeenCalled()
    await act(async () => root?.unmount())
    root = null
    expect(f.client.entriesReady).toHaveBeenLastCalledWith(false)
    expect(f.listeners.size).toBe(0)
  })
  it('waits for onboarding and preferences before connecting or reserving space', async () => {
    const f = await fixture(false)
    expect(vi.mocked(useWorkspaceConnection).mock.calls.at(-1)![0].enabled).toBe(false)
    await f.readPrefs()
    expect(f.client.layout).not.toHaveBeenCalled()
    expect(f.props().workspaceToggle).toBeUndefined()
    await act(async () => f.props().onOnboardingFinished?.())
    expect(vi.mocked(useWorkspaceConnection).mock.calls.at(-1)![0].enabled).toBe(true)
    expect(f.client.layout).toHaveBeenLastCalledWith(null)
  })

  it('keeps the same controller and releases space on collapse without disposing the host', async () => {
    const f = await fixture(true, true)
    await f.readPrefs()
    expect(f.listeners.size).toBe(2)
    expect(f.client.layout).toHaveBeenLastCalledWith(null)
    await act(async () => f.props().workspaceToggle!.toggle())
    expect(f.client.layout).toHaveBeenLastCalledWith({ side: 'left', width: 360, open: true })
    const controller = f.props().workspace!.controller
    await act(async () => f.props().workspace!.setOpen(false))
    expect(f.client.layout).toHaveBeenLastCalledWith(null)
    expect(f.props().workspace!.controller).toBe(controller)
    expect(f.panel.dispose).not.toHaveBeenCalled()
    expect(f.listeners.size).toBe(2)
    await act(async () => root!.unmount())
    root = null
    expect(f.listeners.size).toBe(0)
  })

  it('lets a newer preference broadcast win over a late read, and honors sidebar hiding', async () => {
    const f = await fixture()
    await f.prefs({ ...DEFAULT_AI_PANEL_PREFS, side: 'right', spellcheck: false })
    await f.readPrefs()
    await act(async () => f.props().workspaceToggle!.toggle())
    expect(f.client.layout).toHaveBeenLastCalledWith({ side: 'right', width: 360, open: true })
    await f.prefs({ ...DEFAULT_AI_PANEL_PREFS, showAiSidebar: false })
    expect(f.client.layout).toHaveBeenLastCalledWith(null)
    await act(async () => f.props().workspaceToggle!.toggle())
    expect(f.props().workspace!.layout.panelMode).toBe('open')
  })
})
