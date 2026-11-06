/** @vitest-environment jsdom */
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { defaultAiSettings, type AiSettings } from '@genoffice/ai-provider'
import type { SubscriptionConnection } from '@genoffice/ui'
import type { WorkspacePanelApi, WorkspacePanelEvent } from '../src/shared/workspace-panel-api'
import type { WorkspacePanelClient } from '../src/renderer/src/workspace-ai/panel-client'
import { connectWorkspacePanel } from '../src/renderer/src/workspace-ai/panel-client'
import { WorkspacePanelController } from '../src/renderer/src/workspace-ai/panel-controller'
import { useWorkspaceConnection } from '../src/renderer/src/workspace-ai/use-workspace-connection'
import { makeWorkspaceHost, until } from './helpers/workspace-host-fixture'

vi.mock('../src/renderer/src/workspace-ai/panel-client', () => ({ connectWorkspacePanel: vi.fn() }))
;(
  globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
).IS_REACT_ACT_ENVIRONMENT = true
let root: Root | null = null
afterEach(async () => {
  await act(async () => root?.unmount())
  root = null
  document.body.innerHTML = ''
  vi.clearAllMocks()
  vi.restoreAllMocks()
})

async function fixture(
  state: 'ready' | 'checking' | 'signed-out' = 'ready',
  noticeSeen = true,
  patch: Partial<AiSettings> = {},
) {
  let settings = {
    ...defaultAiSettings(),
    subscriptionNoticeSeen: noticeSeen,
    subscriptionNoticeVersion: noticeSeen ? 2 : 0,
    ...patch,
  }
  const listeners = new Set<(event: WorkspacePanelEvent) => void>()
  const turns: Array<{
    run: Parameters<WorkspacePanelClient['stream']>[0]
    request: Parameters<WorkspacePanelClient['stream']>[1]
  }> = []
  const publish = (
    index: number,
    chunk: Omit<Extract<WorkspacePanelEvent, { kind: 'stream' }>['chunk'], 'requestId'>,
  ) => {
    const turn = turns[index]!
    for (const listener of listeners)
      listener({
        sessionId: 'session',
        kind: 'stream',
        run: turn.run,
        chunk: { ...chunk, requestId: turn.request.requestId },
      })
  }
  const client = {
    subscribe: (listener: (event: WorkspacePanelEvent) => void) => {
      listeners.add(listener)
      return () => {
        listeners.delete(listener)
      }
    },
    stream: vi.fn(async (run, request) => {
      turns.push({ run, request })
    }),
    cancelStream: vi.fn(async (_run, requestId) => {
      const index = turns.findIndex((turn) => turn.request.requestId === requestId)
      if (index >= 0) publish(index, { type: 'done' })
    }),
  } as unknown as WorkspacePanelClient
  let hostFixture!: ReturnType<typeof makeWorkspaceHost>
  let controller!: WorkspacePanelController
  vi.mocked(connectWorkspacePanel).mockImplementation(async (_api, deps) => {
    hostFixture = makeWorkspaceHost({
      ...deps,
      createTransport: (id) => deps.createTransport(id, client),
    })
    hostFixture.addEditor('A')
    controller = new WorkspacePanelController({
      host: hostFixture.host,
      capture: (id) => hostFixture.execution.capture(id),
      refresh: (id) => hostFixture.execution.refresh(id),
      reference: async () => {
        throw new Error('unused')
      },
      create: async () => ({
        origin: { kind: 'workspace', conversationId: crypto.randomUUID() },
        title: '',
      }),
    })
    controller.catalog([{ id: 'A', kind: 'docs', title: 'A', available: true, active: true }], [])
    return {
      host: hostFixture.host,
      controller,
      client,
      dispose: async () => {
        controller.dispose()
        hostFixture.dispose()
      },
    } as Awaited<ReturnType<typeof connectWorkspacePanel>>
  })
  const bridge = {
    getSettings: vi.fn(async () => settings),
    setSettings: vi.fn(async (next: AiSettings) => {
      settings = next
    }),
    openModelSettings: vi.fn(),
  }
  let subscription = {
    statuses: { 'claude-code': { provider: 'claude-code', state, meteredAccepted: false } },
    refresh: vi.fn(),
    busy: false,
  } as unknown as SubscriptionConnection
  subscription.statuses.codex = { provider: 'codex', state: 'signed-out', meteredAccepted: false }
  subscription.statuses.chatgpt = {
    provider: 'chatgpt',
    state: 'signed-out',
    meteredAccepted: false,
  }
  const api = {} as WorkspacePanelApi
  let value!: ReturnType<typeof useWorkspaceConnection>
  function Harness() {
    value = useWorkspaceConnection({
      api,
      bridge,
      subscription,
      lang: 'en',
      unknownError: 'Unknown error',
    })
    return createElement('div', null, value.notices, value.composerControls)
  }
  const container = document.createElement('div')
  document.body.append(container)
  root = createRoot(container)
  await act(async () => {
    root!.render(createElement(Harness))
  })
  await act(async () => {
    await controller.create()
  })
  const p = hostFixture.host.conversations.selectedId!
  const draft = (id: string, text: string) =>
    hostFixture.host.conversations.edit(id, {
      composer: { ...hostFixture.host.conversations.state(id).composer, text },
    })
  return {
    get value() {
      return value
    },
    get host() {
      return hostFixture.host
    },
    controller,
    turns,
    publish,
    draft,
    p,
    bridge,
    client,
    listeners,
    async enableCodex() {
      subscription = {
        ...subscription,
        statuses: {
          ...subscription.statuses,
          codex: { provider: 'codex', state: 'ready', meteredAccepted: false },
        },
      }
      await act(async () => {
        root!.render(createElement(Harness))
      })
    },
    async status(state: 'ready' | 'signed-out') {
      subscription = {
        ...subscription,
        statuses: {
          ...subscription.statuses,
          'claude-code': { provider: 'claude-code', state, meteredAccepted: false },
        },
      }
      await act(async () => {
        root!.render(createElement(Harness))
      })
    },
  }
}

describe('workspace connection and shared send gate', () => {
  it('continues the failed request only after a click, preserving it over later cancelled queue entries', async () => {
    const f = await fixture()
    await act(async () => {
      f.draft(f.p, 'Original request A')
      await f.controller.send(f.p)
      await until(() => f.turns.length === 1)
      f.draft(f.p, 'Cancelled request B')
      await f.controller.send(f.p)
      const queued = f.host.conversations
        .state(f.p)
        .runs.find((entry) => entry.status === 'queued')!
      await f.host.cancel(queued.run)
      f.publish(0, { type: 'error', errorCode: 'rate-limit', errorProvider: 'claude-code' })
      await until(() => !f.host.busy())
    })
    await f.enableCodex()
    expect(f.turns).toHaveLength(1)
    const button = document.querySelector<HTMLButtonElement>('[data-action="continue-with"]')!
    expect(button).not.toBeNull()
    await act(async () => {
      button.click()
      await until(() => f.turns.length === 2)
    })
    expect(f.turns[1]!.request.messages.at(-1)).toMatchObject({
      text: expect.stringContaining('Original request A'),
    })
    expect(JSON.stringify(f.turns[1]!.request.messages.at(-1))).not.toContain('Cancelled request B')
    expect(f.bridge.setSettings).toHaveBeenLastCalledWith(
      expect.objectContaining({ provider: 'codex' }),
    )
    await act(async () => {
      f.publish(1, { type: 'done' })
      await until(() => !f.host.busy())
    })
  })

  it('sends nothing when the person declines the existing local context warning', async () => {
    const confirm = vi.spyOn(window, 'confirm').mockReturnValue(false)
    const f = await fixture('ready', true, {
      provider: 'custom',
      customEndpoint: 'local',
      providers: {
        ...defaultAiSettings().providers,
        custom: { apiKey: '', baseUrl: 'http://127.0.0.1:1234/v1', model: 'small' },
      },
      customEndpoints: [
        {
          id: 'local',
          name: 'Local',
          apiKey: '',
          baseUrl: 'http://127.0.0.1:1234/v1',
          model: 'small',
          kind: 'local',
          modelProfiles: { small: { contextLength: 1 } },
        },
      ],
    })
    await act(async () => {
      f.draft(f.p, 'Large request')
      await f.controller.send(f.p)
      await until(() => !f.host.busy())
    })
    expect(confirm).toHaveBeenCalledOnce()
    expect(f.client.stream).not.toHaveBeenCalled()
    expect(f.value.connection.state(f.p).error?.kind).toBe('context-limit')
    expect(f.host.conversations.state(f.p).error).toContain('context limit')
  })

  it('removes stream subscriptions when an active host unmounts', async () => {
    const f = await fixture()
    await act(async () => {
      f.draft(f.p, 'Request')
      await f.controller.send(f.p)
      await until(() => f.turns.length === 1)
    })
    expect(f.listeners.size).toBe(2)
    await act(async () => {
      root!.unmount()
      root = null
    })
    expect(f.listeners.size).toBe(0)
  })

  it('drops a closed main connection and reconnects only after an explicit retry', async () => {
    const f = await fixture()
    const calls = vi.mocked(connectWorkspacePanel).mock.calls.length
    await act(async () => {
      for (const listener of [...f.listeners]) listener({ sessionId: 'session', kind: 'closed' })
    })
    expect(f.value.panel).toBeNull()
    expect(f.value.failed).toBe(true)
    expect(connectWorkspacePanel).toHaveBeenCalledTimes(calls)
    expect(f.client.stream).not.toHaveBeenCalled()
    await act(async () => f.value.reconnect())
    expect(connectWorkspacePanel).toHaveBeenCalledTimes(calls + 1)
    expect(f.value.panel).not.toBeNull()
    expect(f.value.failed).toBe(false)
    expect(f.client.stream).not.toHaveBeenCalled()
  })

  it('sends nothing on mount, conversation selection or model navigation', async () => {
    const f = await fixture()
    await act(async () => {
      await f.controller.create()
      f.controller.select(f.p)
    })
    const button = document.querySelector<HTMLButtonElement>('.ai-model-chip')
    await act(async () => button?.click())
    expect(f.client.stream).not.toHaveBeenCalled()
    expect(document.body.textContent).toContain('chosen reference scopes')
  })

  it('resumes the original held conversation after checking while preserving another draft', async () => {
    const f = await fixture('checking')
    await act(async () => {
      f.draft(f.p, 'Original P')
      await f.controller.send(f.p)
    })
    expect(f.turns).toHaveLength(0)
    await act(async () => {
      await f.controller.create()
    })
    const q = f.host.conversations.selectedId!
    await act(async () => f.draft(q, 'Draft Q'))
    await f.status('ready')
    await act(async () => {
      await until(() => f.turns.length === 1)
    })
    expect(f.turns[0]!.run.conversationId).toBe(f.p)
    expect(f.turns[0]!.request.messages.at(-1)).toMatchObject({
      text: expect.stringContaining('Original P'),
    })
    expect(f.host.conversations.state(f.p).composer.text).toBe('')
    expect(f.host.conversations.state(q).composer.text).toBe('Draft Q')
    await act(async () => {
      f.publish(0, { type: 'delta', text: 'Answer P' })
      f.publish(0, { type: 'done' })
      await until(() => !f.host.busy())
    })
    expect(f.host.conversations.state(f.p).messages.at(-1)?.text).toBe('Answer P')
    expect(f.host.conversations.state(q).messages).toEqual([])
  })

  it('keeps a disconnected draft and never sends automatically after sign-in', async () => {
    const f = await fixture('signed-out')
    await act(async () => {
      f.draft(f.p, 'Keep this')
      await f.controller.send(f.p)
    })
    await f.status('ready')
    expect(f.turns).toHaveLength(0)
    expect(f.host.conversations.state(f.p).composer.text).toBe('Keep this')
  })

  it('requires the first notice acknowledgement and writes it through the existing settings store', async () => {
    const f = await fixture('ready', false)
    await act(async () => {
      f.draft(f.p, 'Read this')
      await f.controller.send(f.p)
    })
    expect(f.turns).toHaveLength(0)
    const ack = document.querySelector<HTMLButtonElement>('.ai-sub-first button')!
    expect(ack).not.toBeNull()
    await act(async () => {
      ack.click()
      await until(() => f.turns.length === 1)
    })
    expect(f.bridge.setSettings).toHaveBeenCalledOnce()
    await act(async () => {
      f.publish(0, { type: 'done' })
      await until(() => !f.host.busy())
    })
  })

  it('keeps structured errors, usage and truncation in the originating conversation', async () => {
    const f = await fixture()
    await act(async () => {
      f.draft(f.p, 'Request P')
      await f.controller.send(f.p)
      await until(() => f.turns.length === 1)
    })
    await act(async () => {
      await f.controller.create()
    })
    const q = f.host.conversations.selectedId!
    await act(async () => {
      f.publish(0, { type: 'usage', usage: { inputTokens: 17, outputTokens: 5 } })
      f.publish(0, { type: 'input-truncated' })
      f.publish(0, {
        type: 'subscription-usage',
        subscriptionUsage: [
          {
            provider: 'claude-code',
            window: 'five_hour',
            utilization: 1,
            resetsAt: 2_000_000_000_000,
          },
        ],
      })
      f.publish(0, {
        type: 'error',
        errorCode: 'rate-limit',
        errorProvider: 'claude-code',
        errorResetsAt: 2_000_000_000_000,
      })
      await until(() => !f.host.busy())
    })
    expect(f.value.connection.state(f.p)).toMatchObject({
      truncated: true,
      error: { kind: 'rate-limit', provider: 'claude-code' },
      usage: { inputTokens: 17, outputTokens: 5 },
    })
    expect(f.value.connection.state(q)).toMatchObject({
      error: null,
      truncated: false,
      usage: null,
    })
    expect(f.turns).toHaveLength(1)
    expect(document.querySelector('.ai-sub-blocked')).toBeNull()
    await act(async () => f.controller.select(f.p))
    expect(document.querySelector('.ai-sub-blocked')?.getAttribute('data-kind')).toBe('rate-limit')
  })
})
