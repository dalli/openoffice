import { describe, expect, it, vi } from 'vitest'
import { createWorkspaceUndoHistory } from '../studio-plugin/workspace-undo.mjs'
import type { AgentPluginHost } from '../studio-plugin/agent-plugin.mjs'

function fixture(capacity = 100) {
  type Entry = { type: string; resources: number; snapshotResourceCount(): number }
  const entry = (type = 'snapshot:plugin:ieum-agent:replace'): Entry => ({
    type,
    resources: 1,
    snapshotResourceCount() {
      return this.resources
    },
  })
  const stack: Entry[] = []
  const redo: Entry[] = []
  const history = {
    undoStack: stack,
    peekUndoTop: () => stack.at(-1),
    peekRedoTop: () => redo.at(-1),
  }
  const context = { isEditable: true, isFormMode: false }
  let executeHistory: ((services: unknown) => void) | undefined
  const undo = vi.fn(() => {
    const current = stack.pop()!
    current.resources = 2
    redo.push(current)
    // Model the studio's snapshot budget, including its first-undo eviction.
    const live = () => [...stack, ...redo].reduce((sum, item) => sum + item.resources, 0)
    while (live() > capacity - 2 && redo.length > 1) redo.shift()
    while (live() > capacity - 2 && stack.length > 1) stack.shift()
  })
  const host = {
    automation: {
      registerCommand: (command: { execute(services: unknown): void }) => {
        executeHistory = command.execute
      },
      getContext: () => context,
      execute: (name: string) => {
        if (name === 'edit:undo') undo()
        else
          executeHistory?.({
            getInputHandler: () => ({ history }),
            wasm: { snapshotCapacity: () => capacity },
          })
        return { ok: true }
      },
    },
  } as unknown as AgentPluginHost
  const access = createWorkspaceUndoHistory(host)
  return { access, stack, redo, undo, context, entry }
}

describe('complete HWP workspace undo history', () => {
  it('reads a checkpoint without snapshots and undoes only the complete run suffix', () => {
    const f = fixture()
    const earlier = f.entry('user-typing')
    f.stack.push(earlier)
    const checkpoint = f.access.capture()
    f.stack.push(f.entry(), f.entry())
    expect(f.undo).not.toHaveBeenCalled()
    expect(f.access.undo(checkpoint, 2)).toBe(true)
    expect(f.undo).toHaveBeenCalledTimes(2)
    expect(f.stack).toEqual([earlier])
  })

  it.each(['evicted', 'foreign', 'read-only', 'form', 'other-history', 'budget'] as const)(
    'rejects %s before the first native undo',
    (condition) => {
      const f = fixture(condition === 'budget' ? 5 : 100)
      const checkpoint = f.access.capture()
      f.stack.push(f.entry(), f.entry(), f.entry())
      if (condition === 'evicted') f.stack.shift()
      else if (condition === 'foreign') f.stack[1]!.type = 'user-typing'
      else if (condition === 'read-only') f.context.isEditable = false
      else if (condition === 'form') f.context.isFormMode = true
      expect(f.access.undo(condition === 'other-history' ? {} : checkpoint, 3)).toBe(false)
      expect(f.undo).not.toHaveBeenCalled()
    },
  )

  it('leaves room for the first undo snapshot at the studio budget limit', () => {
    const f = fixture(5)
    f.stack.push(f.entry())
    const checkpoint = f.access.capture()
    f.stack.push(f.entry(), f.entry())
    expect(f.access.undo(checkpoint, 2)).toBe(true)
    expect(f.undo).toHaveBeenCalledTimes(2)
    expect(f.stack).toHaveLength(0)
  })

  it('fails closed when studio history hooks are unavailable', () => {
    const access = createWorkspaceUndoHistory({
      automation: { execute: () => ({ ok: true }) },
    } as unknown as AgentPluginHost)
    expect(access.capture()).toBeNull()
    expect(access.undo({}, 1)).toBe(false)
  })
})
