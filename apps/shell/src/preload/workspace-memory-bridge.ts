import {
  WORKSPACE_MEMORY_CHANNELS as C,
  type WorkspaceMemoryApi,
  type WorkspaceMemoryRequest,
} from '../shared/workspace-memory-api'

interface MemoryIpc {
  invoke(channel: string, ...args: unknown[]): Promise<unknown>
  on(channel: string, listener: (event: unknown, value: WorkspaceMemoryRequest) => void): unknown
  removeListener(
    channel: string,
    listener: (event: unknown, value: WorkspaceMemoryRequest) => void,
  ): unknown
}

/** Only the home host is exposed, with no arbitrary channel, path or model operation. */
export function workspaceMemoryBridge(ipc: MemoryIpc): WorkspaceMemoryApi {
  return {
    async register(host) {
      let hostId: string | null = null
      const listener = (_event: unknown, request: WorkspaceMemoryRequest) => {
        if (request.hostId !== hostId) return
        let reply: unknown
        try {
          switch (request.kind) {
            case 'load':
              reply = { ok: true, messages: host.load() }
              break
            case 'append':
              host.append(request.message)
              reply = { ok: true }
              break
            case 'release':
              host.release()
              reply = { ok: true }
              break
          }
        } catch {
          reply = { ok: false }
        }
        void ipc.invoke(C.reply, hostId, request.id, reply).catch(() => {})
      }
      ipc.on(C.request, listener)
      try {
        const id = await ipc.invoke(C.register)
        if (typeof id !== 'string') throw new Error('Memory host unavailable')
        hostId = id
      } catch (error) {
        ipc.removeListener(C.request, listener)
        throw error
      }
      return () => {
        const id = hostId
        hostId = null
        ipc.removeListener(C.request, listener)
        host.release()
        void ipc.invoke(C.unregister, id).catch(() => {})
      }
    },
  }
}
