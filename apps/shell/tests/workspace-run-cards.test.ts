/** @vitest-environment jsdom */
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, describe, expect, it } from 'vitest'
import { WorkspaceRunCards } from '../src/renderer/src/workspace-ai/WorkspaceRunCards'
import { makeWorkspaceHost, until, tick } from './helpers/workspace-host-fixture'

;(
  globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
).IS_REACT_ACT_ENVIRONMENT = true
const labels = {
  undo: 'Undo',
  restoring: 'Restoring',
  reverted: 'Reverted',
  refused: 'Unavailable',
}
let root: Root | null = null
let fixture: ReturnType<typeof makeWorkspaceHost> | null = null
let container: HTMLDivElement | null = null
afterEach(async () => {
  await act(async () => {
    fixture?.dispose()
    root?.unmount()
  })
  container?.remove()
  fixture = null
  root = null
})

describe('workspace run cards', () => {
  it('keeps a permission pending while another conversation is visible and answers only the original store', async () => {
    const f = makeWorkspaceHost({
      prepare: async (_input, run, permissions) => ({
        project: null,
        dispose() {},
        skills: [
          {
            id: 'guarded-test-tool',
            systemPrompt: '',
            tools: [
              {
                name: 'confirm_action',
                description: 'Request permission',
                inputSchema: { type: 'object' },
              },
            ],
            executeTool: async () => {
              const answer = await permissions.ask(
                {
                  server: 'local-test',
                  tool: 'action',
                  exposedName: 'confirm_action',
                  input: { document: 'A' },
                  description: 'Fixture tool',
                  destructive: false,
                },
                run.signal,
              )
              return { output: answer, summary: answer, isError: answer === 'deny' }
            },
          },
        ],
      }),
    })
    fixture = f
    f.addEditor('A')
    f.addEditor('B')
    const p = await f.conversation('P', 'A')
    const q = await f.conversation('Q', 'B')
    container = document.createElement('div')
    document.body.append(container)
    root = createRoot(container)
    const render = (id: string) =>
      root!.render(
        createElement(WorkspaceRunCards, { host: f.host, conversationId: id, lang: 'en', labels }),
      )
    let submitted!: NonNullable<Awaited<ReturnType<typeof p.send>>>
    await act(async () => {
      render(p.id)
      submitted = (await p.send())!
      await until(() => f.requests.length === 1)
    })
    await act(async () => {
      const cb = f.callbacks.get(p.id)![0]!
      cb.onToolCall({ id: 'permission-call', name: 'confirm_action', input: {} })
      cb.onDone()
      await until(() => f.host.permissions(p.id).requests().length === 1)
    })
    expect(container.querySelector('[data-mcp-request]')).not.toBeNull()
    await act(async () => {
      f.host.conversations.select(q.id)
      render(q.id)
      await tick()
    })
    expect(container.querySelector('[data-mcp-request]')).toBeNull()
    expect(f.host.permissions(p.id).requests()).toHaveLength(1)
    expect(f.requests).toHaveLength(1)
    await act(async () => {
      render(p.id)
    })
    await act(async () => {
      ;(container!.querySelector('[data-mcp-request] button') as HTMLButtonElement).click()
      await until(() => f.requests.length === 2)
    })
    expect(f.host.permissions(p.id).requests()).toEqual([])
    expect(f.host.permissions(q.id).requests()).toEqual([])
    await act(async () => {
      const cb = f.callbacks.get(p.id)![1]!
      cb.onDelta('Permission answered')
      cb.onDone()
      await submitted.done
    })
    expect(f.host.conversations.state(q.id).messages).toEqual([])
  })
})
