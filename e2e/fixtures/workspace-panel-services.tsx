import { createRoot } from 'react-dom/client'
import { connectWorkspacePanel } from '../../apps/shell/src/renderer/src/workspace-ai/panel-client'
import { WorkspacePanel } from '../../apps/shell/src/renderer/src/workspace-ai/WorkspacePanel'
import { LocaleProvider } from '../../apps/shell/src/renderer/src/locale'
import { aiSendGate } from '../../packages/ui/src/ai-send-gate'
import type { AgentStreamCallbacks, AgentStreamRequest } from '@genoffice/agent-core'
import type { WorkspaceLayout } from '../../apps/shell/src/shared/workspace-layout'

/** Test-only transport. Every document/history/layout operation uses the production shell preload. */
export async function installWorkspacePanelServices() {
  const requests: AgentStreamRequest[] = []
  const callbacks: AgentStreamCallbacks[] = []
  let cancellations = 0
  const gate = aiSendGate(
    { provider: 'auto', providers: {}, subscriptionNoticeSeen: true },
    {
      'claude-code': { provider: 'claude-code', state: 'ready', meteredAccepted: false },
    },
  )
  const connected = await connectWorkspacePanel(window.aiOfficeWorkspacePanel!, {
    gate: () => gate,
    admit: () => gate.kind === 'run',
    prepare: async () => ({ project: null, skills: [], dispose() {} }),
    errorText: String,
    createTransport: () => ({
      stream(request, cb) {
        requests.push(request)
        callbacks.push(cb)
        return {
          cancel() {
            cancellations++
            queueMicrotask(() => cb.onDone())
          },
        }
      },
    }),
  })
  const container = document.createElement('div')
  document.body.append(container)
  const root = createRoot(container)
  const render = (layout: WorkspaceLayout) =>
    root.render(
      <LocaleProvider initial="en">
        <WorkspacePanel
          mount={{
            controller: connected.controller,
            layout,
            uiScale: 1,
            attachmentBridge: connected.client,
            userSkills: connected.userSkills,
            userMcp: connected.userMcp,
            projectBridge: (
              window as unknown as {
                projectFolderApi: import('@genoffice/project-store').ProjectFolderApi
              }
            ).projectFolderApi,
            setOpen(open) {
              void connected.client.layout({ side: 'right', width: 400, open }).then(render)
            },
          }}
        />
      </LocaleProvider>,
    )
  const off = connected.client.subscribe((event) => {
    if (event.kind === 'layout') render(event.layout)
  })
  render(await connected.client.layout({ side: 'right', width: 400, open: true }))
  return {
    ...connected,
    requests,
    callbacks,
    cancellations: () => cancellations,
    selected: () => connected.host.conversations.selectedId!,
    async dispose() {
      off()
      root.unmount()
      container.remove()
      await connected.dispose()
    },
  }
}

export type WorkspacePanelServicesFixture = Awaited<
  ReturnType<typeof installWorkspacePanelServices>
>
