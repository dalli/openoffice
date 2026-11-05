import { afterEach, describe, expect, it, vi } from 'vitest'
import {
  WorkspacePanelController,
  type WorkspacePanelDeps,
} from '../src/renderer/src/workspace-ai/panel-controller'
import { makeWorkspaceHost, until } from './helpers/workspace-host-fixture'

const cleanups: Array<() => void> = []
afterEach(() => cleanups.splice(0).forEach((off) => off()))
const source = (id: string) => ({
  origin: { kind: 'workspace' as const, conversationId: id },
  title: id,
})
function fixture(patch: Partial<WorkspacePanelDeps> = {}) {
  const f = makeWorkspaceHost()
  f.addEditor('A')
  f.addEditor('B')
  let next = 0
  const capture = vi.fn((id: string) => f.execution.capture(id))
  const reference = vi.fn(async (id: string, scope: 'document' | 'selection') => {
    const { id: refId, binding } = await f.execution.readReference(id, { scope })
    return { id: refId, binding }
  })
  const controller = new WorkspacePanelController({
    host: f.host,
    capture,
    reference,
    create: async () => source(`new-${++next}`),
    ...patch,
  })
  controller.catalog(
    [
      { id: 'A', title: 'Duplicate', kind: 'docs', active: true, available: true },
      { id: 'B', title: 'Duplicate', kind: 'docs', available: true },
      { id: 'unloaded', title: 'Unloaded recovery tab', kind: 'docs', available: false },
    ],
    [source('P'), source('Q')],
  )
  cleanups.push(() => {
    controller.dispose()
    f.dispose()
  })
  return { ...f, controller, capture, reference }
}

describe('workspace panel selection ownership', () => {
  it('keeps a newly created source unique when the main catalog refresh arrives before creation finishes', async () => {
    const f = fixture({
      create: async () => {
        f.controller.catalog(f.controller.documentList(), [
          ...f.controller.sourceList(),
          source('new'),
        ])
        return source('new')
      },
    })
    await f.controller.create()
    expect(
      f.controller
        .sourceList()
        .filter((item) => item.origin.kind === 'workspace' && item.origin.conversationId === 'new'),
    ).toHaveLength(1)
  })

  it('opens source histories without a model call or a document read; only new conversations propose the focused target', async () => {
    const f = fixture()
    await f.controller.open(source('P'))
    const p = f.host.conversations.selectedId!
    expect(f.controller.selection(p).target).toBeNull()
    expect(f.capture).not.toHaveBeenCalled()
    await f.controller.create()
    const q = f.host.conversations.selectedId!
    expect(f.controller.selection(q).target?.binding.document.tabId).toBe('A')
    f.controller.select(p)
    expect(f.host.conversations.selectedId).toBe(p)
    expect(f.controller.selection(p).target).toBeNull()
    expect(f.reference).not.toHaveBeenCalled()
    expect(f.requests).toEqual([])
  })

  it('keeps draft/selection/reference/scroll state per conversation and never follows document focus', async () => {
    const f = fixture()
    await f.controller.create()
    const p = f.host.conversations.selectedId!
    await f.controller.addReference(p, 'B', 'document')
    f.host.conversations.edit(p, {
      composer: { ...f.host.conversations.state(p).composer, text: 'draft P' },
      scrollTop: 130,
    })
    await f.controller.open(source('Q'))
    const q = f.host.conversations.selectedId!
    await f.controller.target(q, 'B')
    f.controller.catalog(
      f.controller.documentList().map((doc) => ({ ...doc, active: doc.id === 'B' })),
      f.controller.sourceList(),
    )
    f.controller.select(p)
    expect(f.host.conversations.state(p).composer.text).toBe('draft P')
    expect(f.host.conversations.state(p).scrollTop).toBe(130)
    expect(f.controller.selection(p).target?.binding.document.tabId).toBe('A')
    expect(f.controller.selection(p).references).toMatchObject([
      { scope: 'document', binding: { document: { tabId: 'B' } } },
    ])
    expect(f.host.conversations.state(q).composer.references).toEqual([])
    expect(f.reference).toHaveBeenCalledTimes(1)
    expect(f.requests).toEqual([])
  })

  it('rejects unavailable recovery tabs without capturing or loading them and retains the draft when a target closes', async () => {
    const f = fixture()
    await f.controller.create()
    const id = f.host.conversations.selectedId!
    await f.controller.target(id, 'unloaded')
    expect(f.controller.selection(id).target).toBeNull()
    expect(f.host.conversations.state(id).composer.target).toBeNull()
    expect(f.capture).toHaveBeenCalledTimes(1)
    expect(f.controller.selection(id).error).toBe(true)
    f.host.conversations.edit(id, {
      composer: { ...f.host.conversations.state(id).composer, text: 'keep this' },
    })
    f.controller.catalog(
      f.controller.documentList().filter((doc) => doc.id !== 'A'),
      f.controller.sourceList(),
    )
    expect(f.controller.canSend(id)).toBe(false)
    await f.controller.send(id)
    expect(f.host.conversations.state(id).composer.text).toBe('keep this')
    expect(f.requests).toEqual([])
  })

  it('does not begin another read or history operation after the window owner disposes the panel controller', async () => {
    const f = fixture()
    await f.controller.create()
    const id = f.host.conversations.selectedId!
    f.capture.mockClear()
    f.load.mockClear()
    f.controller.dispose()
    await f.controller.target(id, 'B')
    await f.controller.addReference(id, 'B', 'document')
    await f.controller.open(source('Q'))
    await f.controller.create()
    expect(f.capture).not.toHaveBeenCalled()
    expect(f.reference).not.toHaveBeenCalled()
    expect(f.load).not.toHaveBeenCalled()
    expect(f.controller.canSend(id)).toBe(false)
  })

  it('does not let a slow history open replace a more recent explicit selection', async () => {
    const f = fixture()
    await f.controller.open(source('Q'))
    const q = f.host.conversations.selectedId!
    let resolve!: (value: { messages: []; foreign: boolean }) => void
    f.load.mockImplementationOnce(
      () =>
        new Promise((done) => {
          resolve = done
        }),
    )
    const pending = f.controller.open(source('P'))
    await until(() => !!resolve)
    f.controller.select(q)
    resolve({ messages: [], foreign: false })
    await pending
    expect(f.host.conversations.selectedId).toBe(q)
  })

  it('ignores an older target capture and routes a late reference to its original conversation', async () => {
    let resolve!: (value: Awaited<ReturnType<WorkspacePanelDeps['capture']>>) => void
    const f = fixture()
    await f.controller.open(source('P'))
    const p = f.host.conversations.selectedId!
    const a = await f.execution.capture('A')
    f.capture.mockImplementationOnce(
      () =>
        new Promise((done) => {
          resolve = done
        }),
    )
    const old = f.controller.target(p, 'A')
    await f.controller.target(p, 'B')
    resolve(a)
    await old
    expect(f.controller.selection(p).target?.binding.document.tabId).toBe('B')
    let referenceDone!: () => void
    const wait = new Promise<void>((done) => {
      referenceDone = done
    })
    f.reference.mockImplementationOnce(async (id, scope) => {
      await wait
      const { id: refId, binding } = await f.execution.readReference(id, { scope })
      return { id: refId, binding }
    })
    const pending = f.controller.addReference(p, 'A', 'document')
    await until(() => f.reference.mock.calls.length > 0)
    await f.controller.open(source('Q'))
    const q = f.host.conversations.selectedId!
    referenceDone()
    await pending
    expect(f.controller.selection(p).references).toHaveLength(1)
    expect(f.controller.selection(q).references).toEqual([])
    expect(f.controller.selection(p).pending).toBe(0)
    expect(f.requests).toEqual([])
  })

  it('uses only the requested reference scope and keeps failed reads out of the composer', async () => {
    const reference = vi.fn(async () => {
      throw new Error('Selection is not available')
    })
    const f = fixture({ reference })
    await f.controller.create()
    const id = f.host.conversations.selectedId!
    await f.controller.addReference(id, 'B', 'selection')
    expect(reference).toHaveBeenCalledWith(expect.any(String), 'selection')
    expect(f.controller.selection(id).references).toEqual([])
    expect(f.host.conversations.state(id).composer.references).toEqual([])
    expect(f.controller.selection(id).error).toBe(true)
  })

  it('sends through the existing guarded host, preserves later drafts and prevents double submission during admission', async () => {
    const f = fixture()
    await f.controller.create()
    const p = f.host.conversations.selectedId!
    f.host.conversations.edit(p, {
      composer: { ...f.host.conversations.state(p).composer, text: 'edit A' },
    })
    const sending = f.controller.send(p)
    const duplicate = f.controller.send(p)
    f.host.conversations.edit(p, {
      composer: { ...f.host.conversations.state(p).composer, text: 'later draft' },
    })
    await Promise.all([sending, duplicate])
    await until(() => f.requests.length === 1)
    expect(f.host.conversations.state(p).composer.text).toBe('later draft')
    await f.controller.open(source('Q'))
    const q = f.host.conversations.selectedId!
    f.callbacks.get(p)![0]!.onDelta('answer for P')
    f.callbacks.get(p)![0]!.onDone()
    await until(() => !f.host.busy())
    expect(f.host.conversations.state(p).messages.at(-1)?.text).toBe('answer for P')
    expect(f.host.conversations.state(q).messages).toEqual([])
  })

  it('keeps a blocked draft and never sends automatically after the gate becomes ready', async () => {
    const f = fixture()
    await f.controller.create()
    const id = f.host.conversations.selectedId!
    f.host.conversations.edit(id, {
      composer: { ...f.host.conversations.state(id).composer, text: 'keep this draft' },
    })
    f.setGate({ kind: 'connect', choice: { connected: false, reason: 'no-subscription' } })
    await f.controller.send(id)
    expect(f.host.conversations.state(id).composer.text).toBe('keep this draft')
    expect(f.requests).toEqual([])
  })
})
