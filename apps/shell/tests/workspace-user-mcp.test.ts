import { afterEach, describe, expect, it, vi } from 'vitest'
import type { McpToolsBridgeLike, OutboundBridgeLike } from '@genoffice/agent-core'
import { WorkspaceUserMcp } from '../src/renderer/src/workspace-ai/user-mcp'
import type { WorkspaceRunHostDeps } from '../src/renderer/src/workspace-ai/run-host'
import { deferred, makeWorkspaceHost, until } from './helpers/workspace-host-fixture'

const cleanups: Array<() => void> = []
afterEach(() => cleanups.splice(0).forEach((off) => off()))
const server = (name: string, permission = 'ask') => ({
  name,
  status: 'ready',
  newToolCount: 0,
  keepResultsInside: true,
  instructions: `Approved ${name} instructions`,
  tools: [
    {
      name: 'lookup',
      exposedName: `mcp__${name}__lookup`,
      description: `Approved ${name} description`,
      inputSchema: { type: 'object', properties: { query: { type: 'string' } } },
      permission,
      pending: false,
    },
  ],
})
function fixture(patch: Partial<WorkspaceRunHostDeps> = {}) {
  let servers = [server('alpha'), server('beta')]
  let internet = true
  const changes = new Set<() => void>()
  const internetChanges = new Set<(on: boolean) => void>()
  const bridge: McpToolsBridgeLike = {
    list: vi.fn(async () => structuredClone(servers)),
    call: vi.fn(async () => ({ ok: true, output: 'PRIVATE_MCP_CONTENT', isError: false })),
    cancel: vi.fn(async () => {}),
    retry: vi.fn(async () => {}),
    openSettings: vi.fn(async () => {}),
    onChanged: (listener) => {
      changes.add(listener)
      return () => void changes.delete(listener)
    },
  }
  const outbound: OutboundBridgeLike = {
    internetToolsEnabled: vi.fn(async () => internet),
    confirmAddresses: vi.fn(async () => true),
    onInternetToolsChanged: (listener) => {
      internetChanges.add(listener)
      return () => void internetChanges.delete(listener)
    },
  }
  const f = makeWorkspaceHost({ prepare: (_input, run) => mcp.prepare(run), ...patch })
  const mcp = new WorkspaceUserMcp(bridge, outbound, f.host, () => 'en')
  f.addEditor('A')
  f.addEditor('B')
  cleanups.push(() => {
    mcp.dispose()
    f.dispose()
  })
  const finish = (id: string, text = 'Done') => {
    const cb = f.callbacks.get(id)!.at(-1)!
    cb.onDelta(text)
    cb.onDone()
  }
  const call = (id: string, name: string, input = {}) => {
    const cb = f.callbacks.get(id)!.at(-1)!
    cb.onToolCall({ id: crypto.randomUUID(), name, input })
    cb.onDone()
  }
  const allow = (id: string, answer: 'once' | 'conversation' | 'deny' = 'once') =>
    f.host.answerPermission(id, f.host.permissions(id).requests()[0]!.id, answer)
  return {
    ...f,
    mcp,
    bridge,
    outbound,
    call,
    finish,
    allow,
    changes,
    internetChanges,
    setInternet(on: boolean) {
      internet = on
      internetChanges.forEach((listener) => listener(on))
    },
    setServers(value: typeof servers) {
      servers = value
      changes.forEach((listener) => listener())
    },
  }
}

describe('workspace MCP with conversation-owned choices and the existing guards', () => {
  it('lists without calling, keeps server choices separate and rereads a queued conversation', async () => {
    const f = fixture()
    const p = await f.conversation('P', 'A')
    const q = await f.conversation('Q', 'B')
    f.mcp.state(p.id).setServerOffInChat('alpha', true)
    await f.mcp.state(q.id).refresh()
    expect(f.requests).toEqual([])
    expect(f.bridge.call).not.toHaveBeenCalled()
    expect(f.host.conversations.state(p.id).mcp.disabledServers).toEqual(['alpha'])
    expect(f.host.conversations.state(q.id).mcp.disabledServers).toEqual([])
    const first = (await p.send())!
    await until(() => f.requests.length === 1)
    const second = (await q.send())!
    f.mcp.state(q.id).setServerOffInChat('beta', true)
    f.finish(p.id)
    await first.done
    await until(() => f.requests.length === 2)
    expect(f.requests[0]!.request.tools.map((tool) => tool.name)).not.toContain(
      'mcp__alpha__lookup',
    )
    expect(f.requests[0]!.request.tools.map((tool) => tool.name)).toContain('mcp__beta__lookup')
    expect(f.requests[1]!.request.tools.map((tool) => tool.name)).toContain('mcp__alpha__lookup')
    expect(f.requests[1]!.request.tools.map((tool) => tool.name)).not.toContain('mcp__beta__lookup')
    f.finish(q.id)
    await second.done
  })

  it('keeps a pending permission when switching and scopes conversation approval to its owner', async () => {
    const f = fixture()
    const p = await f.conversation('P', 'A')
    const q = await f.conversation('Q', 'B')
    const first = (await p.send())!
    await until(() => f.requests.length === 1)
    f.call(p.id, 'mcp__alpha__lookup')
    await until(() => f.host.permissions(p.id).requests().length === 1)
    f.host.conversations.select(q.id)
    expect(f.host.permissions(p.id).requests()).toHaveLength(1)
    expect(f.host.permissions(q.id).requests()).toEqual([])
    expect(f.bridge.call).not.toHaveBeenCalled()
    f.allow(p.id, 'conversation')
    await until(() => f.requests.length === 2)
    expect(f.host.conversations.state(p.id).mcp.alwaysAllowedTools).toEqual(['mcp__alpha__lookup'])
    f.finish(p.id)
    await first.done
    const again = (await p.send())!
    await until(() => f.requests.length === 3)
    f.call(p.id, 'mcp__alpha__lookup')
    await until(() => f.requests.length === 4)
    expect(f.host.permissions(p.id).requests()).toEqual([])
    f.finish(p.id)
    await again.done
    const other = (await q.send())!
    await until(() => f.requests.length === 5)
    f.call(q.id, 'mcp__alpha__lookup')
    await until(() => f.host.permissions(q.id).requests().length === 1)
    expect(f.bridge.call).toHaveBeenCalledTimes(2)
    f.allow(q.id, 'deny')
    await until(() => f.requests.length === 6)
    f.finish(q.id)
    await other.done
  })

  it('keeps received content inside across runs, even after that setting changes, without affecting another conversation', async () => {
    const f = fixture()
    f.setServers([server('alpha', 'always'), server('beta', 'always')])
    const p = await f.conversation('P', 'A')
    const q = await f.conversation('Q', 'B')
    const first = (await p.send())!
    await until(() => f.requests.length === 1)
    f.call(p.id, 'mcp__alpha__lookup')
    await until(() => f.requests.length === 2)
    f.finish(p.id)
    await first.done
    f.setServers([
      { ...server('alpha', 'always'), keepResultsInside: false },
      server('beta', 'always'),
    ])
    const second = (await p.send())!
    await until(() => f.requests.length === 3)
    f.call(p.id, 'mcp__beta__lookup')
    await until(() => f.requests.length === 4)
    expect(f.bridge.call).toHaveBeenCalledTimes(1)
    expect(JSON.stringify(f.requests[3]!.request)).toContain('set to keep what it returns inside')
    f.finish(p.id)
    await second.done
    const other = (await q.send())!
    await until(() => f.requests.length === 5)
    f.call(q.id, 'mcp__beta__lookup')
    await until(() => f.requests.length === 6)
    expect(f.bridge.call).toHaveBeenCalledTimes(2)
    f.finish(q.id)
    await other.done
  })

  it('restores the keep-inside block from recorded MCP results', async () => {
    const f = fixture({
      load: async () => ({
        foreign: false,
        messages: [
          { role: 'user', seq: 0, ts: '2026-11-05', text: 'Earlier request' },
          {
            role: 'assistant',
            seq: 1,
            ts: '2026-11-05',
            text: 'Earlier answer',
            tools: [{ name: 'mcp__alpha__lookup', summary: 'alpha', output: 'Restored content' }],
          },
        ],
      }),
    })
    f.setServers([server('alpha', 'always'), server('beta', 'always')])
    const p = await f.conversation('P', 'A')
    const run = (await p.send())!
    await until(() => f.requests.length === 1)
    f.call(p.id, 'mcp__beta__lookup')
    await until(() => f.requests.length === 2)
    expect(f.bridge.call).not.toHaveBeenCalled()
    f.finish(p.id)
    await run.done
  })

  it('guards document tools with outbound inputs as well as MCP tools', async () => {
    const f = fixture()
    f.setServers([server('alpha', 'always')])
    const editor = f.addEditor('Target')
    const p = await f.conversation('P', 'Target')
    const run = (await p.send())!
    await until(() => f.requests.length === 1)
    f.call(p.id, 'mcp__alpha__lookup')
    await until(() => f.requests.length === 2)
    f.call(p.id, 'write', { text: 'Must not edit', url: 'https://example.test/image' })
    await until(() => f.requests.length === 3)
    expect(JSON.stringify(f.requests[2]!.request)).toContain('set to keep what it returns inside')
    expect(editor.body()).toBe('Target UNSAVED')
    f.finish(p.id)
    await run.done
    expect(f.host.rollbackCards(p.id)).toEqual([])
  })

  it('honors the internet switch in tools, instructions and usability, including a mid-run change', async () => {
    const f = fixture()
    f.setInternet(false)
    const p = await f.conversation('P', 'A')
    const run = (await p.send())!
    await until(() => f.requests.length === 1)
    expect(JSON.stringify(f.requests[0]!.request)).not.toContain('Approved alpha')
    expect(f.requests[0]!.request.tools.map((tool) => tool.name)).toEqual(['read', 'write'])
    expect([...f.mcp.usableServers(p.id)]).toEqual([])
    f.finish(p.id)
    await run.done
    f.setInternet(true)
    const second = (await p.send())!
    await until(() => f.requests.length === 2)
    f.call(p.id, 'mcp__alpha__lookup')
    await until(() => f.host.permissions(p.id).requests().length === 1)
    f.setInternet(false)
    f.allow(p.id)
    await until(() => f.requests.length === 3)
    expect(f.bridge.call).not.toHaveBeenCalled()
    f.finish(p.id)
    await second.done
  })

  it('refuses a tool disabled or awaiting review while its permission card is open', async () => {
    const f = fixture()
    const p = await f.conversation('P', 'A')
    const run = (await p.send())!
    await until(() => f.requests.length === 1)
    f.call(p.id, 'mcp__alpha__lookup')
    await until(() => f.host.permissions(p.id).requests().length === 1)
    const changed = server('alpha')
    changed.tools[0]!.pending = true
    changed.tools[0]!.description = 'UNAPPROVED_DESCRIPTION'
    f.setServers([changed])
    await f.mcp.state(p.id).refresh()
    f.allow(p.id)
    await until(() => f.requests.length === 2)
    expect(f.bridge.call).not.toHaveBeenCalled()
    expect(JSON.stringify(f.requests[1]!.request)).not.toContain('UNAPPROVED_DESCRIPTION')
    f.finish(p.id)
    await run.done
  })

  it('denies pending cards on stop, cancels active calls and discards their late results', async () => {
    const f = fixture()
    const p = await f.conversation('P', 'A')
    const first = (await p.send())!
    await until(() => f.requests.length === 1)
    f.call(p.id, 'mcp__alpha__lookup')
    await until(() => f.host.permissions(p.id).requests().length === 1)
    await f.host.cancel(first.run)
    await first.done
    expect(f.host.permissions(p.id).requests()).toEqual([])
    expect(f.bridge.call).not.toHaveBeenCalled()
    const waiting = deferred()
    vi.mocked(f.bridge.call).mockImplementation(async () => {
      await waiting.promise
      return { ok: true, output: 'LATE_PRIVATE_CONTENT', isError: false }
    })
    const second = (await p.send())!
    await until(() => f.requests.length === 2)
    f.call(p.id, 'mcp__alpha__lookup')
    await until(() => f.host.permissions(p.id).requests().length === 1)
    f.allow(p.id)
    await until(() => vi.mocked(f.bridge.call).mock.calls.length === 1)
    await f.host.cancel(second.run)
    await second.done
    await until(() => vi.mocked(f.bridge.cancel).mock.calls.length === 1)
    waiting.resolve()
    expect(f.requests).toHaveLength(2)
    expect(JSON.stringify(f.append.mock.calls)).not.toContain('LATE_PRIVATE_CONTENT')
    f.mcp.dispose()
    expect(f.changes.size).toBe(0)
    expect(f.internetChanges.size).toBe(0)
  })

  it('does not call after a delayed address approval outlives its run', async () => {
    const f = fixture()
    const waiting = deferred()
    vi.mocked(f.outbound.confirmAddresses).mockImplementation(async () => {
      await waiting.promise
      return true
    })
    const p = await f.conversation('P', 'A')
    const run = (await p.send())!
    f.setServers([server('alpha', 'always')])
    await until(() => f.requests.length === 1)
    f.call(p.id, 'mcp__alpha__lookup', { url: 'https://example.test/private' })
    await until(() => vi.mocked(f.outbound.confirmAddresses).mock.calls.length === 1)
    await f.host.cancel(run.run)
    await run.done
    waiting.resolve()
    await new Promise((resolve) => setTimeout(resolve, 0))
    expect(f.bridge.call).not.toHaveBeenCalled()
    expect(f.requests).toHaveLength(1)
  })

  it('does not start a server call if its final internet check finishes after cancellation', async () => {
    const f = fixture()
    f.setServers([server('alpha', 'always')])
    const p = await f.conversation('P', 'A')
    const run = (await p.send())!
    await until(() => f.requests.length === 1)
    const waiting = deferred()
    let checks = 0
    vi.mocked(f.outbound.internetToolsEnabled).mockImplementation(async () => {
      if (++checks === 3) await waiting.promise
      return true
    })
    f.call(p.id, 'mcp__alpha__lookup')
    await until(() => checks === 3)
    await f.host.cancel(run.run)
    await run.done
    waiting.resolve()
    await new Promise((resolve) => setTimeout(resolve, 0))
    expect(f.bridge.call).not.toHaveBeenCalled()
    expect(f.requests).toHaveLength(1)
  })
})
