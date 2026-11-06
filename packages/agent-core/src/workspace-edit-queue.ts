/** Explicit local UI controls. These are never advertised as model tools. */
export type WorkspaceEditQueueCommand =
  | { kind: 'list' }
  | { kind: 'clear' | 'send'; version: string }
  | { kind: 'remove' | 'focus'; version: string; id: string }
  | { kind: 'update'; version: string; id: string; instruction: string }

export interface WorkspaceEditQueueSnapshot {
  version: string
  items: { id: string; instruction: string; excerpt: string; available: boolean }[]
}

const record = (v: unknown): v is Record<string, unknown> =>
  !!v && typeof v === 'object' && !Array.isArray(v)
const text = (v: unknown, max: number) => typeof v === 'string' && v.length > 0 && v.length <= max

export function isWorkspaceEditQueueCommand(v: unknown): v is WorkspaceEditQueueCommand {
  if (!record(v)) return false
  if (v.kind === 'list') return Object.keys(v).length === 1
  if (!text(v.version, 256)) return false
  if (v.kind === 'send' || v.kind === 'clear') return Object.keys(v).length === 2
  if (!text(v.id, 256)) return false
  if (v.kind === 'focus' || v.kind === 'remove') return Object.keys(v).length === 3
  return (
    v.kind === 'update' &&
    Object.keys(v).length === 4 &&
    text(v.instruction, 500) &&
    (v.instruction as string).trim().length > 0
  )
}

export function isWorkspaceEditQueueSnapshot(v: unknown): v is WorkspaceEditQueueSnapshot {
  return (
    record(v) &&
    text(v.version, 256) &&
    Array.isArray(v.items) &&
    v.items.length <= 10 &&
    v.items.every(
      (i) =>
        record(i) &&
        text(i.id, 256) &&
        text(i.instruction, 500) &&
        typeof i.excerpt === 'string' &&
        i.excerpt.length <= 160 &&
        typeof i.available === 'boolean',
    ) &&
    new Set(v.items.map((i) => i.id)).size === v.items.length
  )
}
