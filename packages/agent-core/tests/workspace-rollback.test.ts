import { describe, expect, it, vi } from 'vitest'
import { createWorkspaceDocumentAdapter } from '../src'

function fixture() {
  let body = 'original'
  let revision = 1
  let readOnly = false
  const capture = vi.fn(() => body)
  const restore = vi.fn((snapshot: unknown) => {
    body = String(snapshot)
    revision++
  })
  const adapter = createWorkspaceDocumentAdapter({
    getState: () => ({
      instanceId: 'doc',
      revision: String(revision),
      selectionId: null,
      readOnly,
    }),
    read: () => ({ text: body, truncated: false }),
    toolAccess: { write: 'write', read: 'read' },
    rollback: { capture, restore },
    createSkill: () => ({
      id: 'doc',
      systemPrompt: '',
      tools: ['read', 'write'].map((name) => ({
        name,
        description: name,
        inputSchema: { type: 'object' },
      })),
      executeTool: (call) => {
        if (call.name === 'read') return { output: body, summary: 'read' }
        body = String(call.input.text)
        revision++
        return { output: 'ok', summary: 'write', mutated: true }
      },
    }),
  })
  return {
    adapter,
    capture,
    restore,
    body: () => body,
    change: () => {
      body = 'manual edit'
      revision++
    },
    readOnly: () => {
      readOnly = true
    },
  }
}
describe('editor-owned workspace rollback', () => {
  it('restores an entire multi-tool run from one private snapshot and consumes the token once', async () => {
    const f = fixture()
    const state = await f.adapter.capture()
    const session = await f.adapter.prepare({ runId: 'run', mode: 'edit', state })
    await session.execute({ id: 'read', name: 'read', input: {} })
    expect(session.rollbackId).toBeNull()
    expect(f.capture).not.toHaveBeenCalled()
    await session.execute({ id: 'one', name: 'write', input: { text: 'first edit' } })
    const id = session.rollbackId!
    await session.execute({ id: 'two', name: 'write', input: { text: 'second edit' } })
    expect(session.rollbackId).toBe(id)
    expect(f.capture).toHaveBeenCalledTimes(1)
    const last = session.state
    session.release()
    await f.adapter.rollback!(id, last)
    expect(f.body()).toBe('original')
    expect(f.restore).toHaveBeenCalledTimes(1)
    await expect(f.adapter.rollback!(id, last)).rejects.toMatchObject({ kind: 'released' })
  })
  it.each(['change', 'readOnly'] as const)(
    'refuses rollback after %s without replacing the document',
    async (action) => {
      const f = fixture()
      const session = await f.adapter.prepare({
        runId: 'run',
        mode: 'edit',
        state: await f.adapter.capture(),
      })
      await session.execute({ id: 'write', name: 'write', input: { text: 'AI edit' } })
      const last = session.state
      session.release()
      f[action]()
      await expect(f.adapter.rollback!(session.rollbackId!, last)).rejects.toMatchObject({
        kind: action === 'change' ? 'stale-document' : 'read-only',
      })
      expect(f.restore).not.toHaveBeenCalled()
    },
  )
})
