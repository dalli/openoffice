import { describe, expect, it, vi } from 'vitest'
import {
  AgentLoop,
  composeSkills,
  type AgentSkill,
  type AgentStreamCallbacks,
  type AgentTransport,
} from '../src'

const tick = () => new Promise((resolve) => setTimeout(resolve, 0))
function deferred<T>() {
  let resolve!: (value: T) => void
  let reject!: (reason: unknown) => void
  const promise = new Promise<T>((ok, fail) => {
    resolve = ok
    reject = fail
  })
  return { promise, resolve, reject }
}
function fixture(verify: NonNullable<AgentSkill['verifyResponse']>) {
  const callbacks: AgentStreamCallbacks[] = []
  const transport: AgentTransport = {
    stream: (_request, cb) => {
      callbacks.push(cb)
      return { cancel: vi.fn() }
    },
  }
  const skill: AgentSkill = {
    id: 'async-check',
    systemPrompt: '',
    tools: [],
    executeTool: () => ({ output: '', summary: '' }),
    verifyResponse: verify,
  }
  const done = vi.fn()
  const error = vi.fn()
  const loop = new AgentLoop({ transport, skill, events: { onDone: done, onError: error } })
  return { loop, callbacks, done, error }
}

describe('async response verification', () => {
  it('waits for the editor and sends a single corrective turn only after its answer', async () => {
    const check = deferred<string | null>()
    const f = fixture(() => check.promise)
    f.loop.run('question')
    await tick()
    f.callbacks[0]!.onDelta('claim')
    f.callbacks[0]!.onDone()
    await tick()
    expect(f.done).not.toHaveBeenCalled()
    expect(f.callbacks).toHaveLength(1)
    check.resolve('Perform the missing edit')
    await tick()
    expect(f.callbacks).toHaveLength(2)
    expect(f.loop.messages.at(-1)).toEqual({ role: 'user', text: 'Perform the missing edit' })
    f.callbacks[1]!.onDelta('corrected')
    f.callbacks[1]!.onDone()
    await tick()
    expect(f.done).toHaveBeenCalledTimes(1)
  })

  it.each(['resolve', 'reject'] as const)(
    'ignores a stale verification %s after cancel and another run',
    async (outcome) => {
      const check = deferred<string | null>()
      const verify = vi.fn().mockReturnValueOnce(check.promise).mockReturnValue(null)
      const f = fixture(verify)
      f.loop.run('old')
      await tick()
      f.callbacks[0]!.onDelta('old claim')
      f.callbacks[0]!.onDone()
      await tick()
      f.loop.cancel()
      f.loop.run('new')
      await tick()
      if (outcome === 'resolve') check.resolve('old correction')
      else check.reject(new Error('old rejection'))
      await tick()
      expect(f.callbacks).toHaveLength(2)
      expect(f.error).not.toHaveBeenCalled()
      expect(f.loop.busy).toBe(true)
      f.callbacks[1]!.onDelta('new answer')
      f.callbacks[1]!.onDone()
      await tick()
      expect(f.loop.messages.some((m) => m.role === 'user' && m.text === 'old correction')).toBe(
        false,
      )
    },
  )

  it('composes synchronous and asynchronous verifiers in order without skipping the later guards', async () => {
    const first = deferred<string | null>()
    const second = vi.fn(() => 'correction')
    const base: AgentSkill = {
      id: 'one',
      systemPrompt: '',
      tools: [],
      executeTool: () => ({ output: '', summary: '' }),
      verifyResponse: () => first.promise,
    }
    const composed = composeSkills('composed', '', [
      base,
      { ...base, id: 'two', verifyResponse: second },
    ])
    const answer = composed.verifyResponse!('claim', [])
    expect(second).not.toHaveBeenCalled()
    first.resolve(null)
    await expect(answer).resolves.toBe('correction')
    expect(second).toHaveBeenCalledTimes(1)
  })
})
