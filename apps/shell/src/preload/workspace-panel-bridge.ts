import {
  WORKSPACE_PANEL_CHANNELS as C,
  type WorkspacePanelApi,
  type WorkspacePanelEvent,
  type WorkspacePanelRequest,
  type WorkspacePanelReply,
  type WorkspacePanelValues,
} from '../shared/workspace-panel-api'

interface PanelIpc {
  invoke(channel: string, ...args: unknown[]): Promise<unknown>
  on(channel: string, listener: (event: unknown, value: WorkspacePanelEvent) => void): unknown
  removeListener(
    channel: string,
    listener: (event: unknown, value: WorkspacePanelEvent) => void,
  ): unknown
}

/** No generic channel or editor API is exposed to the shell page. */
export function workspacePanelBridge(ipc: PanelIpc): WorkspacePanelApi {
  return {
    connect: () => ipc.invoke(C.connect) as ReturnType<WorkspacePanelApi['connect']>,
    request: <R extends WorkspacePanelRequest>(sessionId: string, request: R) =>
      ipc.invoke(C.request, sessionId, request) as Promise<
        WorkspacePanelReply<WorkspacePanelValues[R['kind']]>
      >,
    onEvent: (handler) => {
      const listener = (_event: unknown, value: WorkspacePanelEvent) => handler(value)
      ipc.on(C.event, listener)
      return () => {
        ipc.removeListener(C.event, listener)
      }
    },
  }
}
