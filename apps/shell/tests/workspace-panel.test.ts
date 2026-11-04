/** @vitest-environment jsdom */
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import {
  WorkspacePanel,
  type WorkspacePanelMount,
} from '../src/renderer/src/workspace-ai/WorkspacePanel'
import { WorkspacePanelController } from '../src/renderer/src/workspace-ai/panel-controller'
import { LocaleProvider } from '../src/renderer/src/locale'
import { workspaceLayout } from '../src/shared/workspace-layout'
import { workspacePanelStrings } from '../src/renderer/src/i18n/strings-workspace-panel'
import { makeWorkspaceHost, until } from './helpers/workspace-host-fixture'

;(
  globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }
).IS_REACT_ACT_ENVIRONMENT = true
let root: Root | null = null
const cleanups: Array<() => void> = []
afterEach(async () => {
  await act(async () => {
    root?.unmount()
    cleanups.splice(0).forEach((off) => off())
  })
  document.body.innerHTML = ''
  root = null
})
async function fixture() {
  const f = makeWorkspaceHost()
  f.addEditor('A')
  f.addEditor('B')
  let next = 0
  const controller = new WorkspacePanelController({
    host: f.host,
    capture: (id) => f.execution.capture(id),
    reference: async (id, scope) => {
      const { id: refId, binding } = await f.execution.readReference(id, { scope })
      return { id: refId, binding }
    },
    create: async () => ({
      origin: { kind: 'workspace', conversationId: `new-${++next}` },
      title: `Conversation ${next}`,
    }),
  })
  controller.catalog(
    [
      { id: 'A', title: 'Same title', kind: 'docs', available: true, active: true, pane: 'left' },
      { id: 'B', title: 'Same title', kind: 'docs', available: true, pane: 'right' },
      { id: 'unloaded', title: 'Recovery tab', kind: 'markdown', available: false },
    ],
    [],
  )
  await controller.create()
  const p = f.host.conversations.selectedId!
  await controller.create()
  const q = f.host.conversations.selectedId!
  controller.select(p)
  const container = document.createElement('div')
  document.body.append(container)
  root = createRoot(container)
  let mount: WorkspacePanelMount = {
    controller,
    layout: workspaceLayout(
      { width: 1280, height: 720 },
      { side: 'right', width: 360, open: true },
      'side',
    ),
    uiScale: 1,
    setOpen: vi.fn(),
  }
  const render = (patch: Partial<WorkspacePanelMount> = {}) => {
    mount = { ...mount, ...patch }
    root!.render(
      createElement(LocaleProvider, {
        initial: 'en',
        children: createElement(WorkspacePanel, { mount }),
      }),
    )
  }
  await act(async () => render())
  cleanups.push(() => {
    controller.dispose()
    f.dispose()
  })
  return { ...f, controller, p, q, render, mount }
}
function select(label: string, value: string) {
  const input = [...document.querySelectorAll('select')].find(
    (node) =>
      node.getAttribute('aria-label') === label ||
      node.closest('label')?.textContent?.startsWith(label),
  )!
  input.value = value
  input.dispatchEvent(new Event('change', { bubbles: true }))
}
function button(label: string): HTMLButtonElement {
  return document.querySelector(`button[aria-label="${label}"]`)!
}

describe('workspace shell panel', () => {
  it('shows one surface with distinct target copies and unavailable tabs; switching conversations restores drafts and scroll', async () => {
    const f = await fixture()
    expect(document.querySelectorAll('aside')).toHaveLength(1)
    const targets = document.querySelector('.workspace-ai-field select') as HTMLSelectElement
    expect([...targets.options].map((option) => option.textContent)).toContain(
      'Same title · DOCS · 1 · Left document',
    )
    expect([...targets.options].map((option) => option.textContent)).toContain(
      'Same title · DOCS · 2 · Right document',
    )
    expect(targets.querySelector<HTMLOptionElement>('option[value="unloaded"]')!.disabled).toBe(
      true,
    )
    await act(async () => {
      f.host.conversations.edit(f.p, {
        composer: { ...f.host.conversations.state(f.p).composer, text: 'draft P' },
      })
      const transcript = document.querySelector('.workspace-ai-transcript')!
      transcript.scrollTop = 170
      transcript.dispatchEvent(new Event('scroll', { bubbles: true }))
    })
    await act(async () => f.controller.select(f.q))
    expect(document.querySelector('textarea')!.value).toBe('')
    await act(async () => f.controller.select(f.p))
    expect(document.querySelector('textarea')!.value).toBe('draft P')
    expect(document.querySelector('.workspace-ai-transcript')!.scrollTop).toBe(170)
    expect(f.requests).toEqual([])
  })

  it('acquires a reference only on Add, retains it in its conversation, and clears only that chip', async () => {
    const f = await fixture()
    const spy = vi.spyOn(f.execution, 'readReference')
    await act(async () => {
      select('Reference document', 'B')
      select('Read scope', 'document')
    })
    expect(spy).not.toHaveBeenCalled()
    await act(async () => button('Add reference').click())
    expect(spy).toHaveBeenCalledWith(expect.any(String), { scope: 'document' })
    expect(f.controller.selection(f.p).references).toHaveLength(1)
    await act(async () => f.controller.select(f.q))
    expect(document.querySelector('[aria-label="Remove reference · Same title"]')).toBeNull()
    await act(async () => f.controller.select(f.p))
    await act(async () => button('Remove reference · Same title').click())
    expect(f.controller.selection(f.p).references).toEqual([])
    expect(f.requests).toEqual([])
  })

  it('keeps typing available after the target closes and disables both button and Enter submission', async () => {
    const f = await fixture()
    await act(async () => {
      f.host.conversations.edit(f.p, {
        composer: { ...f.host.conversations.state(f.p).composer, text: 'keep this draft' },
      })
      f.controller.catalog(
        f.controller.documentList().filter((doc) => doc.id !== 'A'),
        f.controller.sourceList(),
      )
    })
    expect(button('Send').disabled).toBe(true)
    const input = document.querySelector('textarea')!
    expect(input.disabled).toBe(false)
    await act(async () =>
      input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true })),
    )
    expect(f.requests).toEqual([])
    expect(input.value).toBe('keep this draft')
    expect(document.body.textContent).toContain('Document closed. Select it again.')
  })

  it('keeps a run alive through collapse and allows the rail to stop its original conversation', async () => {
    const f = await fixture()
    await act(async () => {
      f.host.conversations.edit(f.p, {
        composer: { ...f.host.conversations.state(f.p).composer, text: 'edit A' },
      })
    })
    await act(async () => {
      button('Send').click()
    })
    await act(async () => {
      await until(() => f.requests.length === 1)
    })
    await act(async () =>
      f.render({
        layout: workspaceLayout(
          { width: 1280, height: 720 },
          { side: 'left', width: 360, open: false },
          'side',
        ),
      }),
    )
    expect(f.cancellations).not.toHaveBeenCalled()
    expect(document.querySelector('aside')!.style.width).toBe('40px')
    await act(async () => button('Stop · Conversation 1').click())
    await act(async () => {
      await until(() => !f.host.busy())
    })
    expect(f.cancellations).toHaveBeenCalledTimes(1)
    expect(f.host.conversations.state(f.q).runs).toEqual([])
  })
})

it('provides matching keys and translated panel labels for all 21 languages', () => {
  const entries = Object.entries(workspacePanelStrings)
  expect(entries).toHaveLength(21)
  for (const [lang, strings] of entries) {
    expect(Object.keys(strings).sort()).toEqual(Object.keys(workspacePanelStrings.zh).sort())
    expect(Object.values(strings).every((value) => value.trim().length > 0)).toBe(true)
    if (['ko', 'ja', 'zh', 'zh-TW', 'ar', 'he', 'hi', 'th', 'ru'].includes(lang)) {
      for (const [key, value] of Object.entries(strings))
        expect(value).not.toBe(workspacePanelStrings.en[key as keyof typeof strings])
    }
  }
})
