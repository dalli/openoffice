import { vi } from 'vitest'
import {
  createWorkspaceDocumentAdapter,
  WorkspaceAdapterError,
  type AgentSkill,
  type AgentStreamCallbacks,
  type AgentStreamRequest,
  type AgentTransport,
} from '@genoffice/agent-core'
import { hostWorkspaceEditor } from '@genoffice/agent-core/workspace-editor-host'
import type {
  WorkspaceEditorOperation,
  WorkspaceEditorPort,
  WorkspaceEditorRequest,
  WorkspaceEditorValues,
} from '@genoffice/agent-core/workspace-editor-protocol'
import { defaultAiSettings } from '@genoffice/ai-provider'
import { aiSendGate } from '../../../../packages/ui/src/ai-send-gate'
import {
  WorkspaceExecutionController,
  type WorkspaceEditorClient,
} from '../../src/main/workspace-ai/execution-controller'
import {
  WorkspaceRunHost,
  type WorkspaceRunHostDeps,
} from '../../src/renderer/src/workspace-ai/run-host'

export const tick = () => new Promise((resolve) => setTimeout(resolve, 0))
export async function until(check: () => boolean): Promise<void> {
  for (let i = 0; i < 500; i++) {
    if (check()) return
    await tick()
  }
  throw new Error('Workspace fixture did not reach the requested state')
}
export function deferred() {
  let resolve!: () => void
  const promise = new Promise<void>((done) => {
    resolve = done
  })
  return { promise, resolve }
}

export function makeWorkspaceHost(patch: Partial<WorkspaceRunHostDeps> = {}) {
  let contents = 20
  const editors = new Map<string, ReturnType<typeof addEditor>>()
  const callbacks = new Map<string, AgentStreamCallbacks[]>()
  const requests: Array<{ conversation: string; request: AgentStreamRequest }> = []
  const cancellations = vi.fn()
  const execution = new WorkspaceExecutionController({
    owner: { windowId: 1, panelContentsId: 10, panelGeneration: 1 },
    resolve: (id) => {
      const editor = editors.get(id)
      return editor ? { kind: 'docs', client: editor.client } : null
    },
  })
  let gate = aiSendGate(
    { ...defaultAiSettings(), subscriptionNoticeSeen: true },
    { 'claude-code': { provider: 'claude-code', state: 'ready', meteredAccepted: false } },
  )
  const load = vi.fn(async () => ({ messages: [], foreign: false }))
  const append = vi.fn(async () => 'recorded' as const)
  const prepare = vi.fn(async () => ({
    project: null,
    skills: [] as AgentSkill[],
    dispose: vi.fn(),
  }))
  const host = new WorkspaceRunHost({
    execution,
    load,
    append,
    prepare,
    gate: () => gate,
    admit: () => gate.kind === 'run',
    errorText: (error) => (error instanceof Error ? error.message : String(error)),
    createTransport: (id): AgentTransport => ({
      stream: (request, cb) => {
        requests.push({ conversation: id, request })
        callbacks.set(id, [...(callbacks.get(id) ?? []), cb])
        return {
          cancel: () => {
            cancellations()
            queueMicrotask(() => cb.onDone())
          },
        }
      },
    }),
    ...patch,
  })

  function addEditor(id: string) {
    let body = `${id} UNSAVED`
    let revision = 1
    let available = true
    let seq = 0
    const changed = new Set<() => void>()
    const closed = new Set<() => void>()
    let onRequest: ((request: WorkspaceEditorRequest) => void) | null = null
    let onCancel: ((id: string) => void) | null = null
    const waiting = new Map<
      string,
      { resolve(value: unknown): void; reject(error: Error): void; cleanup(): void }
    >()
    const gate = {
      write: null as ReturnType<typeof deferred> | null,
      verify: null as ReturnType<typeof deferred> | null,
    }
    const notify = () => {
      for (const handler of changed) handler()
    }
    const adapter = createWorkspaceDocumentAdapter({
      masking: 'applied',
      toolAccess: { read: 'read', write: 'write' },
      getState: () => ({
        instanceId: id,
        revision: String(revision),
        selectionId: null,
        readOnly: false,
      }),
      subscribeChanges: (handler) => {
        changed.add(handler)
        return () => {
          changed.delete(handler)
        }
      },
      read: () => ({ text: body, truncated: false }),
      rollback: {
        capture: () => body,
        restore: (snapshot) => {
          body = String(snapshot)
          revision++
          notify()
        },
      },
      createSkill: () => ({
        id,
        systemPrompt: `instructions-${id}`,
        buildContext: () => body,
        tools: ['read', 'write'].map((name) => ({
          name,
          description: name,
          inputSchema: { type: 'object', properties: { text: { type: 'string' } } },
        })),
        executeTool: async (call, signal) => {
          if (call.name === 'read') return { output: body, summary: 'read' }
          if (gate.write) await gate.write.promise
          if (signal?.aborted) throw new WorkspaceAdapterError('cancelled')
          body = String(call.input.text)
          revision++
          notify()
          return { output: 'edited', summary: 'write', mutated: true }
        },
        verifyResponse: async () => {
          if (gate.verify) await gate.verify.promise
          return null
        },
      }),
    })
    const port: WorkspaceEditorPort = {
      onRequest: (handler) => {
        onRequest = handler
        return () => {
          onRequest = null
        }
      },
      onCancel: (handler) => {
        onCancel = handler
        return () => {
          onCancel = null
        }
      },
      event: () => {},
      reply: (reply) => {
        const item = waiting.get(reply.id)
        if (!item) return
        waiting.delete(reply.id)
        item.cleanup()
        if (reply.ok) item.resolve(reply.value)
        else
          item.reject(
            new WorkspaceAdapterError(reply.error === 'failed' ? 'not-ready' : reply.error),
          )
      },
    }
    const off = hostWorkspaceEditor(port, adapter)
    const client: WorkspaceEditorClient = {
      contentsId: ++contents,
      get available() {
        return available
      },
      onChanged: (handler) => {
        changed.add(handler)
        return () => {
          changed.delete(handler)
        }
      },
      onClosed: (handler) => {
        closed.add(handler)
        return () => {
          closed.delete(handler)
        }
      },
      request<O extends WorkspaceEditorOperation>(
        operation: O,
        signal?: AbortSignal,
      ): Promise<WorkspaceEditorValues[O['kind']]> {
        return new Promise((resolve, reject) => {
          if (signal?.aborted || !available) {
            reject(new WorkspaceAdapterError('cancelled'))
            return
          }
          const requestId = `${id}-${++seq}`
          const abort = () => {
            onCancel?.(requestId)
            waiting.delete(requestId)
            reject(new WorkspaceAdapterError('cancelled'))
          }
          signal?.addEventListener('abort', abort, { once: true })
          waiting.set(requestId, {
            resolve: (value) => resolve(value as WorkspaceEditorValues[O['kind']]),
            reject,
            cleanup: () => signal?.removeEventListener('abort', abort),
          })
          queueMicrotask(() => {
            if (waiting.has(requestId)) onRequest?.({ id: requestId, operation })
          })
        })
      },
    }
    const editor = {
      client,
      gate,
      body: () => body,
      manual: () => {
        body = `${id} MANUAL`
        revision++
        notify()
      },
      dispose: () => {
        if (!available) return
        available = false
        for (const handler of closed) handler()
        off()
      },
    }
    editors.set(id, editor)
    return editor
  }
  async function conversation(id: string, tab: string, references: string[] = []) {
    const target = await execution.capture(tab)
    const refs = await Promise.all(
      references.map(async (reference) =>
        execution.readReference((await execution.capture(reference)).id, { scope: 'document' }),
      ),
    )
    const conversationId = await host.conversations.open(
      { kind: 'workspace', conversationId: id },
      target.binding,
    )
    host.conversations.edit(conversationId, {
      composer: {
        ...host.conversations.state(conversationId).composer,
        text: `edit ${tab}`,
        references: refs.map((ref) => ref.binding),
      },
    })
    return {
      id: conversationId,
      target,
      refs,
      send: () =>
        host.submit(
          conversationId,
          target.id,
          refs.map((ref) => ref.id),
        ),
    }
  }
  return {
    host,
    execution,
    addEditor,
    conversation,
    callbacks,
    requests,
    cancellations,
    load,
    append,
    prepare,
    setGate: (value: typeof gate) => {
      gate = value
    },
    dispose: () => {
      host.dispose()
      execution.dispose()
      for (const editor of editors.values()) editor.dispose()
    },
  }
}
