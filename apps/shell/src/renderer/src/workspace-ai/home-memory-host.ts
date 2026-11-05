import type { ChatMessage } from '@genoffice/project-store'
import type { WorkspaceMemoryHost } from '../../../shared/workspace-memory-api'

export interface HomeHistoryMessage {
  role: ChatMessage['role']
  text: string
  skills?: ChatMessage['skills']
  history?: ChatMessage
}

/** A lease protects the original home transcript, including while a project chat is on screen. */
export function homeMemoryHost(deps: {
  busy(): boolean
  messages(): readonly HomeHistoryMessage[]
  append(message: ChatMessage): void
  locked(value: boolean): void
}): WorkspaceMemoryHost {
  let claimed = false
  const openedAt = new Date().toISOString()
  return {
    load() {
      if (deps.busy()) throw new Error('Home conversation is busy')
      claimed = true
      deps.locked(true)
      return deps
        .messages()
        .map((m, seq) => ({
          ...m.history,
          role: m.role,
          text: m.text,
          seq,
          ts: m.history?.ts ?? openedAt,
          ...(m.skills ? { skills: m.skills } : {}),
        }))
        .slice(-200)
    },
    append(message) {
      if (!claimed || deps.busy()) throw new Error('Home conversation is not claimed')
      deps.append({
        ...structuredClone(message),
        seq: deps.messages().length,
        ts: new Date().toISOString(),
      })
    },
    release() {
      claimed = false
      deps.locked(false)
    },
  }
}
