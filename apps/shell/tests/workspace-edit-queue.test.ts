/** @vitest-environment jsdom */
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import {
  WorkspaceEditQueue,
  type WorkspaceEditQueueBridge,
} from '../src/renderer/src/workspace-ai/WorkspaceEditQueue'
import { LocaleProvider } from '../src/renderer/src/locale'
import { isWorkspacePanelRequest } from '../src/main/workspace-ai/panel-request'
import type { WorkspacePanelEvent } from '../src/shared/workspace-panel-api'
import type { WorkspaceEditQueueSnapshot } from '@genoffice/agent-core/workspace-editor-protocol'
import { deferred } from './helpers/workspace-host-fixture'

;(
  globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
).IS_REACT_ACT_ENVIRONMENT = true
let root: Root | null = null
afterEach(async () => {
  await act(async () => root?.unmount())
  root = null
  document.body.innerHTML = ''
})
const target = {
  windowId: 1,
  tabId: 'A',
  documentId: 'doc-A',
  generation: 1,
  kind: 'docs' as const,
}
const snapshot = (text = 'Original instruction'): WorkspaceEditQueueSnapshot => ({
  version: 'v1',
  items: [
    { id: 'q1', instruction: text, excerpt: 'Visible passage', available: true },
    { id: 'gone', instruction: 'Deleted target', excerpt: '', available: false },
  ],
})
async function fixture() {
  const listeners = new Set<(event: WorkspacePanelEvent) => void>()
  const editQueue = vi.fn<WorkspaceEditQueueBridge['editQueue']>(async () => snapshot())
  const bridge = {
    editQueue,
    subscribe(handler: (event: WorkspacePanelEvent) => void) {
      listeners.add(handler)
      return () => {
        listeners.delete(handler)
      }
    },
  }
  const container = document.createElement('div')
  document.body.append(container)
  root = createRoot(container)
  const render = async (doc = target, busy = false) => {
    await act(async () => {
      root!.render(
        createElement(
          LocaleProvider,
          { initial: 'en' },
          createElement(WorkspaceEditQueue, {
            key: doc.documentId,
            bridge,
            document: doc,
            lang: 'en',
            busy,
          }),
        ),
      )
    })
  }
  const button = (label: string) =>
    [...container.querySelectorAll<HTMLButtonElement>('button')].find(
      (el) => (el.getAttribute('aria-label') || el.textContent) === label,
    )!
  const click = async (label: string) => {
    await act(async () => button(label).click())
  }
  await render()
  return {
    container,
    editQueue,
    bridge,
    render,
    button,
    click,
    emit: async (tabId: string) => {
      await act(async () => {
        for (const listener of listeners)
          listener({ sessionId: 'session', kind: 'edit-queue-changed', tabId })
      })
    },
  }
}

describe('workspace edit queue cards', () => {
  it('uses the selected target and preserves the queue until an explicit send or confirmed discard', async () => {
    const f = await fixture()
    expect(f.container.textContent).toContain('2/10')
    expect(f.button('Send 1 edits').disabled).toBe(false)
    await act(async () => f.container.querySelector<HTMLButtonElement>('.ai-queue-text')!.click())
    expect(f.editQueue).toHaveBeenCalledWith(target, { kind: 'focus', id: 'q1', version: 'v1' })
    await f.click('Discard')
    expect(f.editQueue.mock.calls.some(([, command]) => command.kind === 'clear')).toBe(false)
    await f.click('Cancel')
    await f.click('Send 1 edits')
    expect(f.editQueue).toHaveBeenCalledWith(target, { kind: 'send', version: 'v1' })
    await f.render(target, true)
    expect(f.button('Send 1 edits').disabled).toBe(true)
    expect(f.button('Edit').disabled).toBe(true)
  })
  it('discards late reads after switching targets and ignores another document notification', async () => {
    const f = await fixture()
    const held = deferred<WorkspaceEditQueueSnapshot | null>()
    f.editQueue.mockImplementationOnce(() => held.promise)
    await f.emit('A')
    const other = { ...target, tabId: 'B', documentId: 'doc-B' }
    f.editQueue.mockResolvedValue(snapshot('B instruction'))
    await f.render(other)
    await act(async () => held.resolve(snapshot('LATE_A')))
    expect(f.container.textContent).toContain('B instruction')
    expect(f.container.textContent).not.toContain('LATE_A')
    const count = f.editQueue.mock.calls.length
    await f.emit('A')
    expect(f.editQueue).toHaveBeenCalledTimes(count)
    await f.click('Delete')
    expect(f.editQueue).toHaveBeenCalledWith(other, { kind: 'remove', id: 'q1', version: 'v1' })
  })
  it('never retries a rejected send, and keeps its queue available for a deliberate retry', async () => {
    const f = await fixture()
    f.editQueue.mockRejectedValueOnce(new Error('stale-document'))
    await f.click('Send 1 edits')
    expect(f.editQueue.mock.calls.filter(([, command]) => command.kind === 'send')).toHaveLength(1)
    expect(f.container.textContent).toContain('Original instruction')
    expect(f.button('Send 1 edits').disabled).toBe(false)
  })
})

describe('private queue request validation', () => {
  it('accepts bounded explicit actions and rejects document substitutions and malformed commands', () => {
    const request = {
      kind: 'edit-queue',
      document: target,
      command: { kind: 'update', id: 'q1', version: 'v1', instruction: 'Polish' },
    }
    expect(isWorkspacePanelRequest(request)).toBe(true)
    for (const command of [
      { kind: 'update', id: 'q1', version: 'v1', instruction: 'x'.repeat(501) },
      { kind: 'update', id: 'q1', version: 'v1', instruction: ' ' },
      { kind: 'send' },
      { kind: 'list', path: '/other.docx' },
      { kind: 'execute', version: 'v1' },
    ]) {
      expect(isWorkspacePanelRequest({ ...request, command })).toBe(false)
    }
    expect(isWorkspacePanelRequest({ ...request, document: { ...target, kind: 'hwp' } })).toBe(
      false,
    )
    expect(isWorkspacePanelRequest({ ...request, document: { ...target, generation: 0 } })).toBe(
      false,
    )
  })
})
