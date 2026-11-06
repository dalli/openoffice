import { describe, expect, it, vi } from 'vitest'
import type {
  AgentStreamCallbacks,
  AgentStreamRequest,
  WorkspaceWriterUpdate,
} from '@genoffice/agent-core'
import { WorkspaceDocumentWriter } from '../src/renderer/src/workspace-ai/doc-writer'

function fixture() {
  const requests: AgentStreamRequest[] = []
  const callbacks: AgentStreamCallbacks[] = []
  const updates: WorkspaceWriterUpdate[] = []
  const writer = new WorkspaceDocumentWriter(() => {})
  const signal = new AbortController()
  const cancel = vi.fn()
  const beforeTurn = vi.fn(async () => {})
  const run = () =>
    writer.run({
      call: {
        id: 'writer',
        name: 'write_document',
        input: { plan: 'A factual report', context: 'Fixture facts' },
      },
      instruction: 'Original request',
      directive: '\nReply in English.',
      signal: signal.signal,
      beforeTurn,
      transport: {
        stream(request, next) {
          requests.push(request)
          callbacks.push(next)
          return { cancel }
        },
      },
      update: async (value) => {
        updates.push(value)
      },
    })
  return { writer, signal, run, requests, callbacks, updates, cancel, beforeTurn }
}

describe('workspace document writer', () => {
  it('refuses late decisions from a previous writer even when the tool call id is reused', async () => {
    const f = fixture()
    const first = f.run()
    await vi.waitFor(() => expect(f.requests).toHaveLength(1))
    f.callbacks[0]!.onDelta('<p>Old partial</p>')
    f.callbacks[0]!.onStopReason?.('max_tokens')
    f.callbacks[0]!.onDone()
    await vi.waitFor(() => expect(f.writer.state?.partial).toBe(true))
    const oldId = f.writer.state!.id
    f.writer.answer(oldId, false)
    await first
    const second = f.run()
    await vi.waitFor(() => expect(f.requests).toHaveLength(2))
    f.callbacks[1]!.onDelta('<p>New partial</p>')
    f.callbacks[1]!.onStopReason?.('max_tokens')
    f.callbacks[1]!.onDone()
    await vi.waitFor(() => expect(f.writer.state?.partial).toBe(true))
    expect(f.writer.state!.id).not.toBe(oldId)
    f.writer.answer(oldId, true)
    f.writer.stopWriting(oldId)
    expect(f.updates.filter((entry) => entry.kind === 'result')).toHaveLength(1)
    f.writer.answer(f.writer.state!.id, false)
    await second
    expect(f.updates.at(-1)).toMatchObject({ kind: 'result', ok: false })
  })
  it('sends a tool-less request with bounded plan/context and streams only through the live editor callback', async () => {
    const f = fixture()
    const done = f.run()
    await vi.waitFor(() => expect(f.requests).toHaveLength(1))
    expect(f.requests[0]!.tools).toEqual([])
    expect(JSON.stringify(f.requests[0])).toContain('Original request')
    expect(f.beforeTurn).toHaveBeenCalledOnce()
    f.callbacks[0]!.onDelta('<p>Full content</p>')
    f.callbacks[0]!.onDone()
    await done
    expect(f.updates).toEqual([
      { kind: 'ready' },
      { kind: 'progress', text: '<p>Full content</p>' },
      { kind: 'result', ok: true, text: '<p>Full content</p>', truncated: false },
    ])
    expect(f.writer.state).toBeNull()
  })

  it.each([true, false])('waits for an explicit partial decision (%s)', async (keep) => {
    const f = fixture()
    const done = f.run()
    await vi.waitFor(() => expect(f.requests).toHaveLength(1))
    f.callbacks[0]!.onDelta('<p>Partial</p>')
    f.callbacks[0]!.onStopReason?.('max_tokens')
    f.callbacks[0]!.onDone()
    await vi.waitFor(() => expect(f.writer.state?.partial).toBe(true))
    expect(f.updates.some((value) => value.kind === 'result')).toBe(false)
    f.writer.answer(f.writer.state!.id, keep)
    await done
    expect(f.updates.at(-1)).toMatchObject({
      kind: 'result',
      ok: keep,
      ...(keep ? { truncated: true } : {}),
    })
  })

  it('stops only the writer stream and offers its partial, then discards on whole-run cancellation', async () => {
    const f = fixture()
    const done = f.run()
    const settled = done.catch((error: unknown) => error)
    await vi.waitFor(() => expect(f.requests).toHaveLength(1))
    f.callbacks[0]!.onDelta('<p>Partial</p>')
    f.writer.stopWriting(f.writer.state!.id)
    await vi.waitFor(() => expect(f.writer.state?.partial).toBe(true))
    expect(f.cancel).toHaveBeenCalled()
    f.signal.abort()
    expect(await settled).toMatchObject({ message: 'cancelled' })
    f.writer.answer('expired-writer', true)
    expect(f.updates.some((value) => value.kind === 'result')).toBe(false)
    expect(f.writer.state).toBeNull()
  })

  it('retries only an empty writer once and reruns the existing pre-turn guard', async () => {
    const f = fixture()
    const done = f.run()
    await vi.waitFor(() => expect(f.requests).toHaveLength(1))
    f.callbacks[0]!.onDone()
    await vi.waitFor(() => expect(f.requests).toHaveLength(2))
    f.callbacks[1]!.onDone()
    await done
    expect(f.beforeTurn).toHaveBeenCalledTimes(2)
    expect(f.updates.at(-1)).toMatchObject({ kind: 'result', ok: false })
  })

  it('does not call a provider when the editor refuses readiness', async () => {
    const writer = new WorkspaceDocumentWriter(() => {})
    const stream = vi.fn()
    await expect(
      writer.run({
        call: { id: 'x', name: 'write_document', input: { plan: 'Report' } },
        instruction: '',
        directive: '',
        signal: new AbortController().signal,
        transport: { stream },
        beforeTurn: async () => {},
        update: async () => {
          throw new Error('stale-document')
        },
      }),
    ).rejects.toThrow('stale-document')
    expect(stream).not.toHaveBeenCalled()
    expect(writer.state).toBeNull()
  })
})
