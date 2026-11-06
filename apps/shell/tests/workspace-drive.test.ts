import { describe, expect, it, vi } from 'vitest'
import { WorkspaceDrive, type WorkspaceDriveServices } from '../src/main/workspace-ai/drive'
import type { WorkspaceStartedRun } from '../src/main/workspace-ai/execution-controller'
import { deferred } from './helpers/workspace-host-fixture'

function fixture() {
  let status = { connected: true, enabled: true, allowed: true }
  const abort = new AbortController()
  const run = {
    binding: {
      bindingId: 1,
      runId: 'run-one',
      conversationId: 'P',
      owner: { windowId: 1, panelContentsId: 10, panelGeneration: 1 },
    },
    signal: abort.signal,
  } as WorkspaceStartedRun
  const services: WorkspaceDriveServices = {
    status: () => status,
    isDrivePath: (path) => path.startsWith('/drive/'),
    search: vi.fn(async () => ({ ok: true, hits: [] })),
    read: vi.fn(async () => ({
      ok: true,
      page: { name: 'Report', text: 'PRIVATE DRIVE', totalChars: 13, offset: 0 },
    })),
    pick: vi.fn(async () => ({ ok: false, reason: 'cancelled' })),
  }
  const drive = new WorkspaceDrive(services)
  drive.bind(run, [])
  return {
    drive,
    run,
    abort,
    services,
    setStatus: (value: Partial<typeof status>) => {
      status = { ...status, ...value }
    },
  }
}

describe('workspace Drive ownership and existing consent service', () => {
  it('does not search or read on bind/status and refuses other owners and conversations', async () => {
    const f = fixture()
    expect(f.drive.check(f.run.binding).enabled).toBe(true)
    expect(f.services.search).not.toHaveBeenCalled()
    expect(f.services.read).not.toHaveBeenCalled()
    for (const patch of [
      { conversationId: 'Q' },
      { runId: 'other' },
      { bindingId: 2 },
      { owner: { ...f.run.binding.owner, panelGeneration: 2 } },
    ])
      await expect(
        f.drive.tool({ ...f.run.binding, ...patch }, 'read', 'report', 0),
      ).rejects.toThrow()
    expect(f.services.read).not.toHaveBeenCalled()
  })
  it.each(['cancel', 'complete', 'disconnect', 'consent', 'disconnected'] as const)(
    'discards a late read after %s',
    async (kind) => {
      const f = fixture()
      const wait = deferred()
      vi.mocked(f.services.read).mockImplementation(async () => {
        await wait.promise
        return { ok: true, page: { text: 'LATE PRIVATE' } }
      })
      const read = f.drive.tool(f.run.binding, 'read', 'report', 0)
      if (kind === 'cancel') f.abort.abort()
      if (kind === 'complete') f.drive.release(f.run.binding)
      if (kind === 'disconnect') f.drive.dispose()
      if (kind === 'consent') f.setStatus({ allowed: false, enabled: false })
      if (kind === 'disconnected') f.setStatus({ connected: false, enabled: false })
      wait.resolve()
      await expect(read).rejects.toThrow()
    },
  )
  it('retains successful read metadata across runs and restores it without restoring the body', async () => {
    const f = fixture()
    await f.drive.tool(f.run.binding, 'read', 'report', 0)
    f.setStatus({ allowed: false, enabled: false })
    expect(() => f.drive.check(f.run.binding)).toThrow()
    f.drive.release(f.run.binding)
    const q = {
      ...f.run,
      binding: { ...f.run.binding, bindingId: 2, conversationId: 'Q', runId: 'run-two' },
    }
    f.drive.bind(q, [])
    expect(f.drive.check(q.binding).allowed).toBe(false)
    f.drive.restore('Q', [
      {
        seq: 0,
        ts: '',
        role: 'assistant',
        text: '',
        tools: [{ name: 'drive_read', summary: 'Report', isError: false }],
      },
    ])
    expect(() => f.drive.check(q.binding)).toThrow()
  })
  it('keeps Drive attachment provenance for the next run and restores it from attachment metadata', () => {
    const f = fixture()
    f.drive.release(f.run.binding)
    f.drive.bind(f.run, ['/drive/attached.pdf'])
    f.drive.release(f.run.binding)
    f.drive.bind(f.run, [])
    f.setStatus({ allowed: false, enabled: false })
    expect(() => f.drive.check(f.run.binding)).toThrow()
    const q = {
      ...f.run,
      binding: { ...f.run.binding, bindingId: 2, conversationId: 'Q', runId: 'run-two' },
    }
    f.drive.bind(q, [])
    f.drive.restore('Q', [
      {
        seq: 0,
        ts: '',
        role: 'user',
        text: '',
        attachments: [{ name: 'attached.pdf', path: '/drive/attached.pdf' }],
      },
    ])
    expect(() => f.drive.check(q.binding)).toThrow()
  })

  it('blocks Drive targets, references and downloaded attachments before any model turn', () => {
    const f = fixture()
    f.drive.release(f.run.binding)
    f.drive.bind(f.run, ['/drive/material.docx'])
    f.setStatus({ allowed: false })
    expect(() => f.drive.check(f.run.binding)).toThrow()
    expect(f.services.read).not.toHaveBeenCalled()
  })
})
