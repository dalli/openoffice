import { useEffect, useMemo, useRef, useState } from 'react'
import { applyAiPanelPrefs, DEFAULT_AI_PANEL_PREFS, type AiPanelPrefs } from '@genoffice/ui'
import { AppFrameView } from '../AppFrame'
import { useI18n } from '../locale'
import { useSubscriptionConnection } from '../subscription-client'
import { useWorkspaceConnection } from './use-workspace-connection'
import { WorkspacePanelReservation } from './panel-reservation'
import type { WorkspaceLayout } from '../../../shared/workspace-layout'
import { WORKSPACE_PANEL_DEFAULT_PX, workspaceUiScale } from '../../../shared/workspace-layout'
import type { WorkspacePanelMount } from './WorkspacePanel'
import '@genoffice/ui/ai-panel-prefs.css'

const OPEN_KEY = 'shell.workspaceAi.open'

/** The normal shell owns one host; tab changes and panel collapse never replace it. */
export function WorkspaceAppFrame({ initialOnboardingSeen }: { initialOnboardingSeen: boolean }) {
  const subscription = useSubscriptionConnection()
  const { lang, t } = useI18n()
  const [ready, setReady] = useState(initialOnboardingSeen)
  const [prefs, setPrefs] = useState<AiPanelPrefs | null>(null)
  const [open, setOpen] = useState(false)
  const [layout, setLayout] = useState<WorkspaceLayout | null>(null)
  const [reservationFailed, setReservationFailed] = useState(false)
  const [uiScale, setUiScale] = useState(1)
  const reservation = useRef<WorkspacePanelReservation | null>(null)
  const bridge = useMemo(
    () => ({
      getSettings: () => window.aiOffice.getAiSettings(),
      setSettings: window.aiOffice.setAiSettings,
      onSettingsChanged: window.aiOffice.onAiSettingsChanged,
      openModelSettings: window.aiOffice.openAiModelSettings,
    }),
    [],
  )
  const connected = useWorkspaceConnection({
    api: window.aiOfficeWorkspacePanel!,
    bridge,
    subscription,
    lang,
    unknownError: t('askUnknownError'),
    enabled: ready && prefs !== null,
  })
  const panel = connected.panel
  const entryOpen = useRef(open)
  entryOpen.current = open

  useEffect(() => {
    if (!panel) return
    let alive = true
    const off = panel.client.subscribe((event) => {
      if (!alive || event.kind !== 'entry') return
      const next = event.entry.kind === 'toggle' ? !entryOpen.current : true
      entryOpen.current = next
      setOpen(next)
      if (next)
        void panel.controller.enter(
          event.capture,
          event.entry,
          event.entryId
            ? (accepted = true, conversationId?: string) =>
                panel.client.acknowledgeEntry(event.entryId!, accepted, conversationId)
            : undefined,
        )
    })
    void panel.client.entriesReady(true).catch(() => {})
    return () => {
      alive = false
      off()
      void panel.client.entriesReady(false).catch(() => {})
    }
  }, [panel])

  useEffect(() => {
    let alive = true
    let pushed = false
    let initialized = false
    const apply = (next: AiPanelPrefs) => {
      if (!alive) return
      applyAiPanelPrefs(next)
      setPrefs(next)
      if (!initialized) {
        initialized = true
        let remembered = false
        try {
          remembered = localStorage.getItem(OPEN_KEY) === '1'
        } catch {
          // Storage restrictions must not prevent opening the window.
        }
        setOpen(next.showAiSidebar && next.openInNewDocs && remembered)
      } else if (!next.showAiSidebar) setOpen(false)
    }
    const off = window.aiOffice.onAiPanelPrefsChanged((next) => {
      pushed = true
      apply(next)
    })
    void window.aiOffice
      .getAiPanelPrefs()
      .then((next) => {
        if (!pushed) apply(next)
      })
      .catch(() => {
        if (!pushed) apply(DEFAULT_AI_PANEL_PREFS)
      })
    const scale = () =>
      setUiScale(
        workspaceUiScale(
          Number.parseFloat(
            getComputedStyle(document.documentElement).getPropertyValue('--ui-scale'),
          ),
        ),
      )
    scale()
    const observer = new MutationObserver(scale)
    observer.observe(document.documentElement, { attributes: true, attributeFilter: ['style'] })
    return () => {
      alive = false
      off()
      observer.disconnect()
    }
  }, [])

  useEffect(() => {
    if (!panel) return
    setReservationFailed(false)
    const owner = new WorkspacePanelReservation(panel.client, setLayout, () => {
      setReservationFailed(true)
      void panel.dispose().catch(() => {})
    })
    reservation.current = owner
    return () => {
      owner.dispose()
      reservation.current = null
      setLayout(null)
    }
  }, [panel])

  useEffect(() => {
    if (!panel || !prefs) return
    // During migration the shared surface starts closed and reserves no second rail.
    // Main temporarily reduces an explicitly opened surface to a rail in small windows.
    reservation.current?.set(
      open
        ? {
            side: prefs.side,
            width: WORKSPACE_PANEL_DEFAULT_PX,
            open: true,
          }
        : null,
    )
  }, [panel, open, prefs])

  const changeOpen = (next: boolean) => {
    setOpen(next)
    if (prefs?.openInNewDocs) {
      try {
        localStorage.setItem(OPEN_KEY, next ? '1' : '0')
      } catch {
        // The live preference remains usable when persistence is unavailable.
      }
    }
  }
  const mount: WorkspacePanelMount | undefined =
    panel && layout
      ? {
          controller: panel.controller,
          layout,
          uiScale,
          setOpen: changeOpen,
          composerControls: connected.composerControls,
          notices: connected.notices,
          projectBridge: window.projectFolderApi,
          attachmentBridge: panel.client,
          userSkills: panel.userSkills,
          userMcp: panel.userMcp,
          editQueueBridge: panel.client,
        }
      : undefined
  return (
    <AppFrameView
      initialOnboardingSeen={initialOnboardingSeen}
      connection={subscription}
      workspace={mount}
      workspaceToggle={
        ready && prefs
          ? {
              open: layout?.panelMode === 'open',
              title: connected.failed || reservationFailed ? t('askUnknownError') : undefined,
              toggle() {
                if (connected.failed || reservationFailed) {
                  setReservationFailed(false)
                  connected.reconnect()
                }
                changeOpen(layout?.panelMode !== 'open')
              },
            }
          : undefined
      }
      onOnboardingFinished={() => setReady(true)}
    />
  )
}
