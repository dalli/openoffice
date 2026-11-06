import { useEffect, useState } from 'react'
import { createRoot } from 'react-dom/client'
import { useSubscriptionConnection } from '@genoffice/ui'
import { LocaleProvider } from '../../apps/shell/src/renderer/src/locale'
import { WorkspacePanel } from '../../apps/shell/src/renderer/src/workspace-ai/WorkspacePanel'
import { useWorkspaceConnection } from '../../apps/shell/src/renderer/src/workspace-ai/use-workspace-connection'
import type { WorkspaceLayout } from '../../apps/shell/src/shared/workspace-layout'

/** Real connection hook and preload; the spec selects only its local fake model endpoint. */
export function installWorkspaceConnection() {
  let current: ReturnType<typeof useWorkspaceConnection> | null = null
  const container = document.createElement('div')
  document.body.append(container)
  const root = createRoot(container)
  const bridge = {
    getSettings: () => window.aiOffice.getAiSettings(),
    setSettings: window.aiOffice.setAiSettings,
    onSettingsChanged: window.aiOffice.onAiSettingsChanged,
  }
  function Harness() {
    const subscription = useSubscriptionConnection({ api: () => window.aiOfficeSubscription })
    const connected = useWorkspaceConnection({
      api: window.aiOfficeWorkspacePanel!,
      bridge,
      subscription,
      lang: 'en',
      unknownError: 'Unknown error',
    })
    current = connected
    const [layout, setLayout] = useState<WorkspaceLayout | null>(null)
    const panel = connected.panel
    useEffect(() => {
      if (!panel) return
      let alive = true
      const off = panel.client.subscribe((event) => {
        if (event.kind === 'layout') setLayout(event.layout)
      })
      void panel.client.layout({ side: 'right', width: 420, open: true }).then((value) => {
        if (alive) setLayout(value)
      })
      return () => {
        alive = false
        off()
      }
    }, [panel])
    return panel && layout ? (
      <WorkspacePanel
        mount={{
          controller: panel.controller,
          layout,
          uiScale: 1,
          composerControls: connected.composerControls,
          notices: connected.notices,
          setOpen: (open) => {
            void panel.client.layout({ side: 'right', width: 420, open }).then(setLayout)
          },
        }}
      />
    ) : null
  }
  root.render(
    <LocaleProvider initial="en">
      <Harness />
    </LocaleProvider>,
  )
  return {
    current: () => current,
    dispose() {
      root.unmount()
      container.remove()
    },
  }
}
