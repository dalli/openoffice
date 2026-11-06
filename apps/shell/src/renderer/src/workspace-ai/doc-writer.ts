import {
  streamText,
  type AgentToolCall,
  type AgentTransport,
  type AgentTokenUsage,
  type WorkspaceWriterUpdate,
} from '@genoffice/agent-core'
import {
  buildDocWriterRequest,
  countFragmentBlocks,
  DOC_MAX_CHARS,
  extractFragment,
} from '../../../../../docs/src/renderer/ai/doc-writer-request'

export interface WorkspaceWritingState {
  id: string
  blocks: number
  partial: boolean
}

/** Panel-owned transport and decision; the editor alone owns the temporary draft and final edit. */
export class WorkspaceDocumentWriter {
  state: WorkspaceWritingState | null = null
  private decide: ((keep: boolean) => void) | null = null
  private stop: AbortController | null = null
  constructor(private readonly changed: () => void) {}
  answer(id: string, keep: boolean): void {
    if (this.state?.id === id && this.state.partial) this.decide?.(keep)
  }
  stopWriting(id: string): void {
    if (this.state?.id === id && !this.state.partial) this.stop?.abort()
  }

  async run(options: {
    call: AgentToolCall
    instruction: string
    directive: string
    transport: AgentTransport & { settled?(): Promise<void> }
    onUsage?(usage: AgentTokenUsage): void
    signal: AbortSignal
    beforeTurn(): Promise<void>
    update(update: WorkspaceWriterUpdate): Promise<void>
  }): Promise<void> {
    if (this.stop) throw new Error('Writing already active')
    const id = crypto.randomUUID()
    const { call, signal } = options
    const spec = {
      plan: String(call.input.plan ?? ''),
      title: String(call.input.title ?? ''),
      context: String(call.input.context ?? ''),
      instruction: options.instruction,
    }
    const { system, user } = buildDocWriterRequest(spec, options.directive)
    this.stop = new AbortController()
    const writingSignal = AbortSignal.any([signal, this.stop.signal])
    let pending: Promise<void> | null = null
    let latest: string | null = null
    let failure: unknown
    const update = (text: string) => {
      this.state = { id, blocks: countFragmentBlocks(text), partial: false }
      this.changed()
      latest = text
      if (!pending)
        pending = (async () => {
          while (latest !== null) {
            const next = latest
            latest = null
            if (signal.aborted) throw new Error('cancelled')
            await options.update({ kind: 'progress', text: next })
          }
        })()
          .catch((error) => {
            failure = error
            this.stop?.abort()
          })
          .finally(() => {
            pending = null
          })
    }
    const cancel = () => this.decide?.(false)
    signal.addEventListener('abort', cancel, { once: true })
    try {
      await options.update({ kind: 'ready' })
      this.state = { id, blocks: 0, partial: false }
      this.changed()
      const attempt = async () => {
        await options.transport.settled?.()
        await options.beforeTurn()
        if (signal.aborted) throw new Error('cancelled')
        return streamText({
          transport: options.transport,
          system,
          user,
          signal: writingSignal,
          maxChars: DOC_MAX_CHARS,
          extract: (raw) => ({ text: extractFragment(raw.slice(0, DOC_MAX_CHARS)) }),
          onProgress: update,
          onUsage: options.onUsage,
        })
      }
      let outcome = await attempt()
      if (outcome.status === 'empty' && !writingSignal.aborted) outcome = await attempt()
      await pending
      if (latest !== null && !signal.aborted && !failure)
        await options.update({ kind: 'progress', text: latest })
      await options.transport.settled?.()
      if (failure) throw failure
      if (signal.aborted) throw new Error('cancelled')
      let keep = outcome.status === 'complete'
      if (outcome.status === 'partial') {
        this.state = { id, blocks: countFragmentBlocks(outcome.text), partial: true }
        keep = await new Promise<boolean>((resolve) => {
          this.decide = resolve
          if (signal.aborted) resolve(false)
          this.changed()
        })
      }
      if (signal.aborted) throw new Error('cancelled')
      await options.update(
        keep && outcome.status !== 'empty'
          ? {
              kind: 'result',
              ok: true,
              text: outcome.text,
              truncated: outcome.status === 'partial',
            }
          : {
              kind: 'result',
              ok: false,
              error:
                outcome.status === 'empty'
                  ? outcome.error
                  : 'the user discarded the partial content',
            },
      )
    } finally {
      signal.removeEventListener('abort', cancel)
      this.stop.abort()
      this.stop = null
      this.decide = null
      this.state = null
      this.changed()
    }
  }
}
