import { EventEmitter } from 'node:events'
import type { WebContents } from 'electron'
import { describe, expect, it } from 'vitest'
import { WorkspaceLegacySources } from '../src/main/workspace-ai/legacy-sources'
import { emitLegacyChatEvent, onLegacyChatEvent } from '../../docs/src/main/legacy-chat-events'

describe('live legacy chat references', () => {
  it('lists only hosts still in this window, releases navigation listeners and retains no transcript', () => {
    const sources = new WorkspaceLegacySources()
    const sender = Object.assign(new EventEmitter(), { id: 1, isDestroyed: () => false })
    const wc = sender as unknown as WebContents
    const ref = { projectId: 'default', chatId: 'unsaved-one' }
    sources.receive({ kind: 'resolved', sender: wc, ref })
    expect(sources.sources(() => undefined)).toEqual([])
    expect(sources.sources(() => 'Untitled')).toEqual([
      { origin: { kind: 'legacy-document', ...ref }, title: 'Untitled' },
    ])
    sources.receive({
      kind: 'resolved',
      sender: wc,
      ref: { ...ref, chatId: 'saved' },
      previous: ref,
    })
    expect(sender.listenerCount('destroyed')).toBe(1)
    sender.emit('did-start-navigation', {}, 'anchor', true, true)
    expect(sources.sources(() => 'Saved')).toHaveLength(1)
    sender.emit('did-start-navigation', {}, 'reload', false, true)
    expect(sources.sources(() => 'Saved')).toEqual([])
    expect(sender.listenerCount('destroyed')).toBe(0)
    sources.receive({ kind: 'resolved', sender: wc, ref })
    sender.emit('render-process-gone')
    expect(sources.sources(() => 'Untitled')).toEqual([])
  })
  it('keeps metadata observers from changing a successful existing storage operation', () => {
    const off = onLegacyChatEvent(() => {
      throw new Error('observer failed')
    })
    try {
      expect(() => emitLegacyChatEvent({ kind: 'renamed' })).not.toThrow()
    } finally {
      off()
    }
  })
})
