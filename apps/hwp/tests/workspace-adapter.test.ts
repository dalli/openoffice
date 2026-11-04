import { describe, expect, it, vi } from 'vitest'
import type { AgentToolCall } from '@genoffice/agent-core'
import type { AgentSelectionResult, AgentWorkspaceVersion } from '../studio-plugin/agent-plugin.mjs'
import type { AgentBridge } from '../src/renderer/ai/agent-bridge'
import { createHwpWorkspaceAdapter } from '../src/renderer/ai/workspace-adapter'

function deferred<T>() {
  let resolve!: (value: T) => void
  return {
    promise: new Promise<T>((done) => {
      resolve = done
    }),
    resolve: (value: T) => resolve(value),
  }
}

function setup(options: { memo?: boolean; selected?: boolean } = {}) {
  let text = 'unsaved source material'
  let revision = 1
  let instance = 'hwp-instance:1'
  let foreign = 0
  let readOnly = false
  let writes = 0
  let selection: AgentSelectionResult = options.selected
    ? {
        ok: true,
        own: false,
        inTable: false,
        contiguous: true,
        segments: [{ sec: 0, para: 0, offset: 8, length: 6, text: 'source' }],
      }
    : { ok: false, reason: 'none' }
  const question = deferred<boolean>()
  const asked = deferred<void>()
  const confirm = vi.fn(() => {
    asked.resolve()
    return question.promise
  })
  const onChange = vi.fn()
  const undoBaseline = vi.fn()
  let beforeInvoke: ((method: string) => Promise<void>) | null = null
  const invoke = vi.fn(async (expected: AgentWorkspaceVersion, method: string, args: unknown[]) => {
    await beforeInvoke?.(method)
    if (expected.instance !== instance || expected.revision !== String(revision)) {
      return { ok: false, reason: 'changed' }
    }
    let value: unknown
    switch (method) {
      case 'readText':
        value = text
        break
      case 'listFields':
        value = []
        break
      case 'listTables':
        value = { count: 0, tables: [] }
        break
      case 'listNotes':
        value = { count: 0, notes: [] }
        break
      case 'listStyles':
        value = { paragraph: [], character: [] }
        break
      case 'listMemos':
        value = { count: 0, memos: [], supported: true }
        break
      case 'readSelectionSnapshot': {
        const segments = args[0] as { offset: number; length: number; text: string }[]
        if (
          segments.some(
            (segment) =>
              text.slice(segment.offset, segment.offset + segment.length) !== segment.text,
          )
        ) {
          throw new Error('the captured selection changed')
        }
        value = segments.map((segment) => segment.text).join('\n')
        break
      }
      case 'replaceText': {
        if (options.memo && !(args[3] as { approvedMemoIds?: number[] })?.approvedMemoIds) {
          value = {
            ok: false,
            reason: 'memo_removal_requires_approval',
            count: 1,
            memos: [{ id: 9 }],
          }
          break
        }
        text = text.replace(String(args[0]), String(args[1]))
        writes++
        revision++
        value = { ok: true, inTable: false, place: { sec: 0, para: 0, offset: 8, length: 6 } }
        break
      }
      case 'replaceSelection': {
        const segments = args[0] as { offset: number; length: number; text: string }[]
        const replacement = (args[1] as string[])[0]!
        const selected = segments[0]!
        text =
          text.slice(0, selected.offset) +
          replacement +
          text.slice(selected.offset + selected.length)
        writes++
        revision++
        const after = { ...selected, text: replacement, length: replacement.length }
        value = { ok: true, segments: [after], place: after }
        selection = { ok: false, reason: 'none' }
        break
      }
      default:
        return { ok: false, reason: 'unsupported-method' }
    }
    return { ok: true, instance, revision: String(revision), foreignChanges: foreign, value }
  })
  const bridge = {
    workspaceState: vi.fn(async () => ({
      instance,
      revision: String(revision),
      text,
      foreignChanges: foreign,
      masking: 'unsupported',
      selection: structuredClone(selection),
    })),
    workspaceInvoke: invoke,
  } as unknown as AgentBridge
  const adapter = createHwpWorkspaceAdapter({
    bridge,
    label: (key) => key,
    onChange,
    confirmMemoRemoval: confirm,
    readOnly: () => readOnly,
    onUndoBaseline: undoBaseline,
  })
  return {
    adapter,
    bridge,
    invoke,
    onChange,
    undoBaseline,
    confirm,
    question,
    asked,
    text: () => text,
    writes: () => writes,
    userEdit(value = ' changed by user') {
      text += value
      revision++
      foreign++
    },
    replaceInstance() {
      instance = 'hwp-instance:2'
    },
    loseSelection() {
      selection = { ok: false, reason: 'none' }
    },
    unsupportedSelection() {
      selection = { ok: false, reason: 'unsupported' }
    },
    setReadOnly() {
      readOnly = true
    },
    delayInvoke(callback: (method: string) => Promise<void>) {
      beforeInvoke = callback
    },
  }
}

const replace = (id = 'call-1'): AgentToolCall => ({
  id,
  name: 'replace_text',
  input: { find: 'source', replace: 'updated' },
})

describe('HWP workspace document adapter', () => {
  it('reads current unsaved content with explicit unsupported masking and truncation metadata', async () => {
    const { adapter } = setup()
    const state = await adapter.capture()
    expect(await adapter.read(state, { scope: 'document' })).toMatchObject({
      text: 'unsaved source material',
      masking: 'unsupported',
      truncated: false,
      state,
    })
    expect(await adapter.read(state, { scope: 'document', maxChars: 7 })).toMatchObject({
      text: 'unsaved',
      truncated: true,
    })
  })

  it('keeps the captured selection when focus leaves and validates its actual current text', async () => {
    const { adapter, loseSelection, invoke } = setup({ selected: true })
    const state = await adapter.capture()
    loseSelection()
    const read = await adapter.read(state, { scope: 'selection' })
    expect(read.text).toBe('source')
    expect(invoke).toHaveBeenCalledWith(
      { instance: state.instanceId, revision: state.revision },
      'readSelectionSnapshot',
      [[expect.objectContaining({ offset: 8, text: 'source' })]],
    )
    const session = await adapter.prepare({ runId: 'p', mode: 'edit', state })
    expect(session.context).toContain('Selected text')
    expect(session.context).toContain('source')
  })

  it('does not expose network, media, file creation, save or reference writes', async () => {
    const { adapter, writes } = setup()
    const state = await adapter.capture()
    const reference = await adapter.prepare({ runId: 'reference', mode: 'reference', state })
    expect(reference.tools.some((tool) => tool.name === 'read_document')).toBe(true)
    for (const name of [
      'replace_text',
      'read_url',
      'list_pictures',
      'analyze_picture',
      'insert_image',
      'generate_image',
      'save',
      'create_document',
    ]) {
      expect(reference.tools.some((tool) => tool.name === name)).toBe(false)
    }
    await expect(reference.execute(replace())).rejects.toMatchObject({ kind: 'unknown-tool' })
    expect(writes()).toBe(0)
    reference.release()
    const edit = await adapter.prepare({ runId: 'edit', mode: 'edit', state })
    expect(edit.tools.some((tool) => tool.name === 'replace_text')).toBe(true)
    expect(edit.tools.some((tool) => tool.name === 'read_url')).toBe(false)
  })

  it('reuses the existing skill change callback and advances revision for one undoable operation', async () => {
    const { adapter, writes, text, onChange, undoBaseline } = setup()
    const session = await adapter.prepare({
      runId: 'p',
      mode: 'edit',
      state: await adapter.capture(),
    })
    const result = await session.execute(replace())
    expect(result.result).toMatchObject({ mutated: true, output: 'Replaced.' })
    expect(result.state.revision).toBe('2')
    expect(text()).toBe('unsaved updated material')
    expect(writes()).toBe(1)
    expect(onChange).toHaveBeenCalledExactlyOnceWith(
      expect.objectContaining({ kind: 'replace', from: 'source', to: 'updated' }),
    )
    expect(undoBaseline).toHaveBeenCalledExactlyOnceWith(0)
    await expect(session.execute(replace())).rejects.toMatchObject({ kind: 'duplicate-call' })
    expect(writes()).toBe(1)
  })

  it('refreshes its frozen selected range after its own successful replacement', async () => {
    const { adapter, text } = setup({ selected: true })
    const state = await adapter.capture()
    const session = await adapter.prepare({ runId: 'p', mode: 'edit', state })
    const { state: after, result } = await session.execute({
      id: 'selected',
      name: 'replace_selection',
      input: { paragraphs: ['replacement'] },
    })
    expect(result.mutated).toBe(true)
    expect(text()).toBe('unsaved replacement material')
    expect((await adapter.read(after, { scope: 'selection' })).text).toBe('replacement')
  })

  it('rejects manual edits or another document generation after capture', async () => {
    for (const action of ['edit', 'swap'] as const) {
      const current = setup()
      const state = await current.adapter.capture()
      if (action === 'edit') current.userEdit()
      else current.replaceInstance()
      await expect(
        current.adapter.prepare({ runId: 'p', mode: 'edit', state }),
      ).rejects.toMatchObject({ kind: 'stale-document' })
      expect(current.writes()).toBe(0)
    }
  })

  it('fails closed for a selection the plugin cannot address instead of widening it to the document', async () => {
    const current = setup()
    current.unsupportedSelection()
    await expect(current.adapter.capture()).rejects.toMatchObject({ kind: 'unsupported-read' })
    expect(current.writes()).toBe(0)
  })

  it('preserves a declined memo approval without retrying the change', async () => {
    const current = setup({ memo: true })
    const session = await current.adapter.prepare({
      runId: 'p',
      mode: 'edit',
      state: await current.adapter.capture(),
    })
    const pending = session.execute(replace())
    await current.asked.promise
    current.question.resolve(false)
    expect((await pending).result.isError).toBe(true)
    expect(current.invoke.mock.calls.filter(([, method]) => method === 'replaceText')).toHaveLength(
      1,
    )
    expect(current.writes()).toBe(0)
    expect(current.onChange).not.toHaveBeenCalled()
  })

  it('checks revision in the plugin after a queued bridge operation rather than only before awaiting it', async () => {
    const current = setup()
    const session = await current.adapter.prepare({
      runId: 'p',
      mode: 'edit',
      state: await current.adapter.capture(),
    })
    const gate = deferred<void>()
    const entered = deferred<void>()
    current.delayInvoke(async (method) => {
      if (method === 'replaceText') {
        entered.resolve()
        await gate.promise
      }
    })
    const pending = session.execute(replace())
    await entered.promise
    current.userEdit()
    gate.resolve()
    await expect(pending).rejects.toMatchObject({ kind: 'stale-document' })
    expect(current.writes()).toBe(0)
  })

  it('rejects a memo approval retry after a manual edit while the question was pending', async () => {
    const current = setup({ memo: true })
    const session = await current.adapter.prepare({
      runId: 'p',
      mode: 'edit',
      state: await current.adapter.capture(),
    })
    const pending = session.execute(replace())
    await current.asked.promise
    current.userEdit()
    current.question.resolve(true)
    await expect(pending).rejects.toMatchObject({ kind: 'stale-document' })
    expect(current.writes()).toBe(0)
    expect(current.onChange).not.toHaveBeenCalled()
  })

  it('preserves memo approval for the exact unchanged operation', async () => {
    const current = setup({ memo: true })
    const session = await current.adapter.prepare({
      runId: 'p',
      mode: 'edit',
      state: await current.adapter.capture(),
    })
    const pending = session.execute(replace())
    await current.asked.promise
    current.question.resolve(true)
    expect((await pending).result.mutated).toBe(true)
    expect(current.writes()).toBe(1)
    expect(current.invoke).toHaveBeenLastCalledWith(expect.anything(), 'replaceText', [
      'source',
      'updated',
      null,
      { approvedMemoIds: [9] },
    ])
  })

  it.each(['abort', 'release', 'read-only'] as const)(
    'does not retry a pending memo edit after %s',
    async (action) => {
      const current = setup({ memo: true })
      const session = await current.adapter.prepare({
        runId: 'p',
        mode: 'edit',
        state: await current.adapter.capture(),
      })
      const controller = new AbortController()
      const pending = session.execute(replace(), controller.signal)
      await current.asked.promise
      if (action === 'abort') controller.abort()
      else if (action === 'release') session.release()
      else current.setReadOnly()
      current.question.resolve(true)
      if (action === 'read-only') expect((await pending).result.isError).toBe(true)
      else
        await expect(pending).rejects.toMatchObject({
          kind: action === 'abort' ? 'cancelled' : expect.any(String),
        })
      expect(current.writes()).toBe(0)
      expect(current.onChange).not.toHaveBeenCalled()
    },
  )
})
