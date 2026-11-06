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
import { deferred, makeWorkspaceHost, until } from './helpers/workspace-host-fixture'
import { WorkspaceUserSkills } from '../src/renderer/src/workspace-ai/user-skills'
import { WorkspaceUserMcp } from '../src/renderer/src/workspace-ai/user-mcp'
import type { ProjectFolderApi } from '@genoffice/project-store'
import type { WorkspaceRunHostDeps } from '../src/renderer/src/workspace-ai/run-host'
import { workspaceDocsActions } from '../src/renderer/src/workspace-ai/docs-actions'

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
async function fixture(patch: Partial<WorkspaceRunHostDeps> = {}) {
  const f = makeWorkspaceHost(patch)
  f.addEditor('A')
  f.addEditor('B')
  let next = 0
  const controller = new WorkspacePanelController({
    host: f.host,
    capture: (id) => f.execution.capture(id),
    refresh: (id) => f.execution.refresh(id),
    reference: async (id, scope) => {
      const { id: refId, binding } = await f.execution.readReference(id, { scope })
      return { id: refId, binding }
    },
    create: async (projectId) => ({
      origin: projectId
        ? { kind: 'project', projectId, chatId: `new-${++next}` }
        : { kind: 'workspace', conversationId: `new-${++next}` },
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
  it('keeps a delayed mind map in its original conversation and enables follow-up drafts only for its original document', async () => {
    const tree = { label: 'MAP_PRIVATE_TOPIC', block: 0, children: [{ label: 'Branch' }] }
    const f = await fixture({
      prepare: async () => ({
        project: null,
        skills: [
          {
            id: 'fixture-map',
            systemPrompt: '',
            tools: [
              {
                name: 'make_map',
                description: 'map',
                inputSchema: { type: 'object', properties: {} },
              },
            ],
            executeTool: () => ({
              output: 'Map rendered',
              summary: 'map',
              display: { kind: 'mindmap', tree },
            }),
          },
        ],
        dispose() {},
      }),
    })
    await act(async () => {
      f.controller.preset(f.p, 'Make a map')
      await f.controller.send(f.p)
      await until(() => f.requests.length === 1)
      f.controller.select(f.q)
      const turn = f.callbacks.get(f.p)![0]
      turn.onToolCall({ id: 'map-result', name: 'make_map', input: {} })
      turn.onDone()
      await until(() => f.requests.length === 2)
      const answer = f.callbacks.get(f.p)![1]
      answer.onDelta('Map ready')
      answer.onDone()
      await until(() => !f.host.busy())
    })
    expect(document.querySelector('.ai-mindmap')).toBeNull()
    expect(f.host.mindMaps(f.q)).toEqual([])
    expect(f.host.mindMaps(f.p)).toHaveLength(1)
    expect(JSON.stringify(f.append.mock.calls)).not.toContain('MAP_PRIVATE_TOPIC')
    expect(JSON.stringify(f.requests[1].request)).not.toContain('display')
    await act(async () => f.controller.select(f.p))
    expect(document.querySelector('.ai-mindmap-title')!.textContent).toBe('Mind map')
    // Citation numbers remain visible; unguarded jumps and document insertion are not offered.
    expect(document.querySelector('.ai-mindmap-badge')!.getAttribute('role')).toBeNull()
    expect(document.querySelectorAll('.ai-mindmap-actions button')).toHaveLength(3)
    await act(async () => {
      document
        .querySelector('.ai-mindmap-viewport')!
        .dispatchEvent(
          new WheelEvent('wheel', { deltaY: 100, clientX: 40, clientY: 40, cancelable: true }),
        )
    })
    const canvas = document.querySelector<HTMLElement>('.ai-mindmap-canvas')!
    const zoom = canvas.style.transform
    expect(zoom).not.toContain('scale(1)')
    await act(async () => {
      f.controller.preset(f.p, 'Continue typing')
    })
    expect(canvas.style.transform).toBe(zoom)
    const node = document.querySelector('.ai-mindmap-node.root')!
    await act(async () => node.dispatchEvent(new MouseEvent('click', { bubbles: true })))
    expect(document.querySelector('textarea')!.value).toContain('MAP_PRIVATE_TOPIC')
    expect(f.requests).toHaveLength(2)
    await act(async () => f.controller.target(f.p, 'B'))
    expect(document.querySelectorAll('.ai-mindmap-node[role="button"]')).toHaveLength(0)
    await act(async () => {
      f.controller.preset(f.p, 'Keep this draft')
      document
        .querySelector('.ai-mindmap-node')!
        .dispatchEvent(new MouseEvent('click', { bubbles: true }))
    })
    expect(document.querySelector('textarea')!.value).toBe('Keep this draft')
    expect(f.requests).toHaveLength(2)
  })
  it('fills each DOCX quick action without sending, keeps it in its conversation, and disables actions for a closed target', async () => {
    const f = await fixture()
    const target = f.controller.selection(f.p).target
    for (const action of workspaceDocsActions('en')) {
      await act(async () => button(action.label).click())
      expect(document.querySelector('textarea')!.value).toBe(action.prompt)
      expect(document.activeElement).toBe(document.querySelector('textarea'))
      expect(f.controller.selection(f.p).target).toEqual(target)
      expect(f.requests).toEqual([])
    }
    await act(async () => f.controller.select(f.q))
    expect(document.querySelector('textarea')!.value).toBe('')
    await act(async () => f.controller.select(f.p))
    expect(document.querySelector('textarea')!.value).toBe(workspaceDocsActions('en')[3].prompt)
    await act(async () => f.controller.catalog([], []))
    expect(button('AI Summarize').disabled).toBe(true)
    expect(f.requests).toEqual([])
  })
  it('keeps one project and MCP control while switching conversations', async () => {
    const f = await fixture()
    const mcp = new WorkspaceUserMcp(
      {
        list: async () => [{ name: 'alpha', status: 'ready', tools: [] }],
        call: vi.fn(),
        cancel: vi.fn(),
        retry: vi.fn(),
        openSettings: vi.fn(),
        onChanged: () => () => {},
      },
      { internetToolsEnabled: async () => true, confirmAddresses: async () => false },
      f.host,
      () => 'en',
    )
    cleanups.push(() => mcp.dispose())
    const projectBridge = {
      listProjects: async () => [],
      getDefaultParentFolder: async () => '/projects',
    } as unknown as ProjectFolderApi
    await act(async () => f.render({ userMcp: mcp, projectBridge }))
    for (const id of [f.q, f.p, f.q, f.p]) {
      await act(async () => f.controller.select(id))
      expect(document.querySelectorAll('.workspace-ai-surface .ai-project-chip')).toHaveLength(1)
      expect(document.querySelectorAll('.workspace-ai-surface .ai-mcp-server-btn')).toHaveLength(1)
    }
    expect(f.requests).toEqual([])
  })

  it('keeps skill chips and edited proposal cards across conversation switches and settles a late save in its original history', async () => {
    const f = await fixture({ prepare: (input, run) => skills.prepare(input, run) })
    const saved = deferred<unknown>()
    const bridge = {
      list: vi.fn(async () => [
        { id: 'app:minutes', name: 'minutes', title: 'minutes', source: 'app', enabled: true },
      ]),
      aiRead: vi.fn(async () => ({
        ok: true,
        skill: { id: 'app:minutes', body: 'PRIVATE_SKILL_BODY', files: [] },
      })),
      aiReadFile: vi.fn(),
      openSettings: vi.fn(async () => {}),
      onChanged: () => () => {},
      suggestName: vi.fn(),
      checkName: vi.fn(async () => ({ ok: true })),
      saveFromCard: vi.fn(() => saved.promise),
    }
    const skills = new WorkspaceUserSkills(bridge, f.host, () => 'en')
    cleanups.push(() => skills.dispose())
    const old = (globalThis as { aiOfficeUserSkills?: unknown }).aiOfficeUserSkills
    Object.assign(globalThis, { aiOfficeUserSkills: bridge })
    cleanups.push(() => Object.assign(globalThis, { aiOfficeUserSkills: old }))
    await act(async () => {
      await skills.refresh()
      f.render({ userSkills: skills })
    })
    await act(async () => skills.composer(f.p).onPick(skills.items()[0]!))
    expect(document.querySelector('.ai-skill-chips')?.textContent).toContain('minutes')
    await act(async () => f.controller.select(f.q))
    expect(document.querySelector('.ai-skill-chips')).toBeNull()
    await act(async () => f.controller.select(f.p))
    expect(document.querySelector('.ai-skill-chips')?.textContent).toContain('minutes')
    await act(async () =>
      f.host.conversations.edit(f.p, {
        composer: {
          ...f.host.conversations.state(f.p).composer,
          text: 'Make meeting notes',
        },
      }),
    )
    await act(async () => f.controller.send(f.p))
    await act(async () => until(() => f.requests.length === 1))
    await act(async () => {
      const cb = f.callbacks.get(f.p)![0]!
      cb.onToolCall({
        id: 'proposal',
        name: 'propose_skill',
        input: { name: 'team-minutes', body: 'Keep the agenda', description: 'Meeting notes' },
      })
      cb.onDone()
      await until(() => skills.proposals(f.p).length === 1)
    })
    const body = document.querySelector<HTMLTextAreaElement>('.ai-skill-card-body')!
    expect(body).toBeTruthy()
    await act(async () => {
      Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')!.set!.call(
        body,
        'Edited repeatable instructions',
      )
      body.dispatchEvent(new Event('input', { bubbles: true }))
    })
    await act(async () => f.controller.select(f.q))
    expect(document.querySelector('.ai-skill-card')).toBeNull()
    await act(async () => f.controller.select(f.p))
    expect(document.querySelector<HTMLTextAreaElement>('.ai-skill-card-body')!.value).toBe(
      'Edited repeatable instructions',
    )
    await act(async () => new Promise((resolve) => setTimeout(resolve, 300)))
    const save = [...document.querySelectorAll('button')].find((b) => b.textContent === 'Save')!
    expect(save.disabled).toBe(false)
    await act(async () => save.click())
    await act(async () => f.controller.select(f.q))
    await act(async () => f.controller.select(f.p))
    expect(document.querySelector<HTMLFieldSetElement>('.workspace-ai-skill-save')!.disabled).toBe(
      true,
    )
    expect(bridge.saveFromCard).toHaveBeenCalledTimes(1)
    await act(async () => f.controller.select(f.q))
    await act(async () => saved.resolve({ ok: true, id: 'app:team-minutes' }))
    expect(skills.proposals(f.p)[0]!.status).toBe('saved')
    expect(skills.proposals(f.q)).toEqual([])
    expect(f.append).toHaveBeenCalledWith(
      f.host.conversations.state(f.p).origin,
      expect.objectContaining({ tools: [expect.objectContaining({ name: 'propose_skill' })] }),
    )
    expect(bridge.saveFromCard).toHaveBeenCalledWith(
      expect.objectContaining({ body: 'Edited repeatable instructions' }),
    )
  })

  it('keeps late attachment picks, chips, removal and rejection notices in the original conversation', async () => {
    const f = await fixture()
    const waiting = deferred<import('../../docs/src/shared/ipc').AttachmentAddResult>()
    const bridge = {
      pickAttachments: vi.fn(() => waiting.promise),
      addAttachments: vi.fn(),
      pasteAttachment: vi.fn(),
    }
    await act(async () => f.render({ attachmentBridge: bridge }))
    await act(async () => button('Attach local files (or drag them onto the panel)').click())
    await act(async () => f.controller.select(f.q))
    await act(async () =>
      waiting.resolve({
        accepted: [{ path: '/one.md', name: 'one.md', ext: 'md', sizeBytes: 10 }],
        rejected: ['Unreadable file'],
      }),
    )
    expect(bridge.pickAttachments).toHaveBeenCalledWith(f.p)
    expect(f.host.conversations.state(f.q).composer.attachments).toEqual([])
    expect(document.querySelector('.workspace-ai-attachments')!.textContent).not.toContain('one.md')
    expect(document.querySelector('.workspace-ai-attachments')!.textContent).not.toContain(
      'Unreadable file',
    )
    await act(async () => f.controller.select(f.p))
    expect(document.querySelector('.workspace-ai-attachments')!.textContent).toContain('one.md')
    expect(document.querySelector('[role="alert"]')!.textContent).toBe('Unreadable file')
    await act(async () => button('Remove attachment · one.md').click())
    expect(f.host.conversations.state(f.p).composer.attachments).toEqual([])
    expect(f.requests).toEqual([])
  })

  it('keeps a late Google picker result and translated refusal in the conversation that opened it', async () => {
    const f = await fixture()
    const waiting = deferred<import('../src/shared/drive-api').DriveAttachResult>()
    const bridge = {
      pickAttachments: vi.fn(),
      addAttachments: vi.fn(async () => ({
        accepted: [{ path: '/drive.md', name: 'drive.md', ext: 'md', sizeBytes: 10 }],
        rejected: ['Unreadable copy'],
      })),
      pasteAttachment: vi.fn(),
      driveStatus: vi.fn(async () => ({ connected: true, enabled: true, allowed: true })),
      pickDriveAttachments: vi.fn(() => waiting.promise),
    }
    await act(async () => f.render({ attachmentBridge: bridge }))
    await act(async () => button('Attach from Google Drive').click())
    expect(button('Attach local files (or drag them onto the panel)').disabled).toBe(true)
    await act(async () => f.controller.select(f.q))
    await act(async () =>
      waiting.resolve({ ok: true, paths: ['/drive.md'], skipped: 1, failed: 0 }),
    )
    expect(bridge.pickDriveAttachments).toHaveBeenCalledWith(f.p)
    expect(bridge.addAttachments).toHaveBeenCalledWith(f.p, ['/drive.md'])
    expect(f.host.conversations.state(f.q).composer.attachments).toEqual([])
    await act(async () => f.controller.select(f.p))
    expect(document.querySelector('.workspace-ai-attachments')!.textContent).toContain('drive.md')
    expect(document.querySelector('[role="alert"]')!.textContent).toContain('Not attached: 1')
    expect(document.querySelector('[role="alert"]')!.textContent).toContain('Unreadable copy')
    expect(f.requests).toEqual([])
  })

  it('caps attachment chips with a translated notice and routes dropped files to the selected conversation', async () => {
    const f = await fixture()
    const files = Array.from({ length: 21 }, (_, index) => ({
      path: `/file-${index}.md`,
      name: `file-${index}.md`,
      ext: 'md',
      sizeBytes: 10,
    }))
    const bridge = {
      pickAttachments: vi.fn(async () => ({ accepted: files, rejected: [] })),
      addAttachments: vi.fn(async () => ({ accepted: [files[0]!], rejected: [] })),
      pasteAttachment: vi.fn(),
    }
    await act(async () => f.render({ attachmentBridge: bridge }))
    await act(async () => button('Attach local files (or drag them onto the panel)').click())
    expect(f.host.conversations.state(f.p).composer.attachments).toHaveLength(20)
    expect(document.querySelector('[role="alert"]')!.textContent).toBe(
      'Attach up to 20 files per message.',
    )
    await act(async () => f.controller.select(f.q))
    const old = window.aiOffice
    Object.assign(window, { aiOffice: { pathForFile: () => '/file-0.md' } })
    const drop = new Event('drop', { bubbles: true, cancelable: true })
    Object.defineProperty(drop, 'dataTransfer', {
      value: { types: ['Files'], files: [new File(['content'], 'file-0.md')] },
    })
    await act(async () => document.querySelector('aside')!.dispatchEvent(drop))
    Object.assign(window, { aiOffice: old })
    expect(bridge.addAttachments).toHaveBeenCalledWith(f.q, ['/file-0.md'])
    expect(f.host.conversations.state(f.q).composer.attachments).toEqual([files[0]])
    expect(f.requests).toEqual([])
  })

  it('keeps the project for a new conversation and renders notices only in their original conversation', async () => {
    const f = await fixture()
    await act(async () => f.controller.create('chosen-project'))
    const original = f.host.conversations.selectedId!
    await act(async () =>
      f.controller.projectNotices(original, [{ kind: 'instructions-changed', files: ['agents'] }]),
    )
    expect(document.querySelector('[data-notice="instructions-changed"]')).not.toBeNull()
    await act(async () => button('New conversation').click())
    const next = f.host.conversations.selectedId!
    expect(next).not.toBe(original)
    expect(f.host.conversations.state(next).origin).toMatchObject({
      kind: 'project',
      projectId: 'chosen-project',
    })
    expect(document.querySelector('[data-notice="instructions-changed"]')).toBeNull()
    await act(async () => f.controller.select(original))
    expect(document.querySelector('[data-notice="instructions-changed"]')).not.toBeNull()
    expect(f.requests).toEqual([])
  })
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
