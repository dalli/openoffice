import type { AgentPluginHost } from './agent-plugin.mjs'

export declare function createWorkspaceUndoHistory(host: AgentPluginHost): {
  capture(): unknown
  undo(checkpoint: unknown, steps: number): boolean
}
