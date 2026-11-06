import { describe, expect, it, vi } from 'vitest'
import { WorkspacePanelReservation } from '../src/renderer/src/workspace-ai/panel-reservation'
import type { WorkspacePanelEvent } from '../src/shared/workspace-panel-api'
import { workspaceLayout, type WorkspacePanelLayout } from '../src/shared/workspace-layout'
import { until } from './helpers/workspace-host-fixture'

function deferred<T>() {
  let resolve!: (value: T) => void
  const promise = new Promise<T>((done) => {
    resolve = done
  })
  return { promise, resolve }
}

const right: WorkspacePanelLayout = { side: 'right', width: 360, open: true }
const layout = workspaceLayout({ width: 1280, height: 720 }, right)

function fixture() {
  const listeners = new Set<(event: WorkspacePanelEvent) => void>()
  const change = vi.fn()
  const failed = vi.fn()
  const client = {
    layout: vi.fn(async (_panel: WorkspacePanelLayout | null) => layout),
    subscribe(listener: (event: WorkspacePanelEvent) => void) {
      listeners.add(listener)
      return () => {
        listeners.delete(listener)
      }
    },
  }
  const owner = new WorkspacePanelReservation(client, change, failed)
  const emit = (event: WorkspacePanelEvent) => {
    for (const listener of listeners) listener(event)
  }
  return { owner, client, change, failed, listeners, emit }
}

describe('shell workspace reservation', () => {
  it('waits for main, serializes writes and ignores a superseded acknowledgement', async () => {
    const f = fixture()
    const first = deferred<typeof layout>()
    const second = deferred<typeof layout>()
    f.client.layout.mockReturnValueOnce(first.promise).mockReturnValueOnce(second.promise)
    f.owner.set(right)
    await until(() => f.client.layout.mock.calls.length === 1)
    f.owner.set(null)
    expect(f.change).not.toHaveBeenCalled()
    expect(f.client.layout).toHaveBeenCalledOnce()
    first.resolve(layout)
    await until(() => f.client.layout.mock.calls.length === 2)
    expect(f.change).not.toHaveBeenCalled()
    const hidden = workspaceLayout({ width: 1280, height: 720 })
    second.resolve(hidden)
    await until(() => f.change.mock.calls.length === 1)
    expect(f.change).toHaveBeenLastCalledWith(hidden)
    expect(f.client.layout).toHaveBeenLastCalledWith(null)
    f.owner.dispose()
  })

  it('follows main resize, scale, split and presentation events without changing preferences', async () => {
    const f = fixture()
    f.owner.set(right)
    await until(() => f.change.mock.calls.length === 1)
    for (const next of [
      workspaceLayout({ width: 1024, height: 768, uiScale: 1.5 }, right, 'side'),
      workspaceLayout({ width: 1600, height: 900 }, right, 'stack'),
      workspaceLayout({ width: 1600, height: 900 }),
    ]) {
      f.emit({ sessionId: 'session', kind: 'layout', layout: next })
      expect(f.change).toHaveBeenLastCalledWith(next)
    }
    expect(f.client.layout).toHaveBeenCalledOnce()
    f.owner.dispose()
  })

  it('keeps a resize pushed before the reservation acknowledgement', async () => {
    const f = fixture()
    const pending = deferred<typeof layout>()
    f.client.layout.mockReturnValueOnce(pending.promise)
    f.owner.set(right)
    await until(() => f.client.layout.mock.calls.length === 1)
    const resized = workspaceLayout({ width: 900, height: 720 }, right, 'side')
    f.emit({ sessionId: 'session', kind: 'layout', layout: resized })
    expect(f.change).not.toHaveBeenCalled()
    pending.resolve(layout)
    await until(() => f.change.mock.calls.length === 1)
    expect(f.change).toHaveBeenLastCalledWith(resized)
    f.owner.dispose()
  })

  it.each(['closed', 'unmounted'] as const)(
    'discards queued writes and late replies after %s',
    async (reason) => {
      const f = fixture()
      const pending = deferred<typeof layout>()
      f.client.layout.mockReturnValueOnce(pending.promise)
      f.owner.set(right)
      await until(() => f.client.layout.mock.calls.length === 1)
      f.owner.set(null)
      if (reason === 'closed') f.emit({ sessionId: 'session', kind: 'closed' })
      else f.owner.dispose()
      f.change.mockClear()
      pending.resolve(layout)
      await Promise.resolve()
      await Promise.resolve()
      await Promise.resolve()
      expect(f.change).not.toHaveBeenCalled()
      expect(f.client.layout).toHaveBeenCalledOnce()
      expect(f.listeners.size).toBe(0)
    },
  )

  it('releases a failed reservation and reports it once', async () => {
    const f = fixture()
    f.client.layout.mockRejectedValueOnce(new Error('closed'))
    f.owner.set(right)
    await until(() => f.failed.mock.calls.length === 1)
    expect(f.change).toHaveBeenLastCalledWith(null)
    expect(f.listeners.size).toBe(0)
    f.owner.set(null)
    await Promise.resolve()
    expect(f.client.layout).toHaveBeenCalledOnce()
  })
})
