import { useEffect, useRef, useState, useSyncExternalStore } from 'react'
import type { AiSettings } from '@genoffice/ai-provider/browser'
import { localModelCapabilities } from '@genoffice/ai-provider/browser'
import type { Lang } from '@genoffice/i18n'
import {
  AiBlockedNotice,
  AiConnectPrompt,
  AiFirstConnectionNotice,
  AiInputTruncatedNotice,
  AiModelPicker,
  AiSubscriptionUsageLine,
  AiToolsUnsupportedNotice,
  AiUsageChip,
  useAiSendGate,
  type AiModelPickerBridge,
  type SubscriptionConnection,
} from '@genoffice/ui'
import type { WorkspacePanelApi } from '../../../shared/workspace-panel-api'
import { connectWorkspacePanel } from './panel-client'
import { WorkspaceConnection } from './connection'
import type { WorkspacePendingSend } from './run-host'
import { askBlockedRunOf } from '../kb-ask-connection'
import { workspacePanelText } from '../i18n/strings-workspace-panel'
import { runErrorCauseOf, shouldRecheckAfter } from '../../../../../docs/src/renderer/ai/run-state'

type Connected = Awaited<ReturnType<typeof connectWorkspacePanel>>
const idleSubscribe = () => () => {}
const idleVersion = () => 0

/** Opt-in production services for AppFrame.workspace; mounting never starts a model turn. */
export function useWorkspaceConnection({
  api,
  bridge,
  subscription,
  lang,
  unknownError,
  enabled = true,
}: {
  api: WorkspacePanelApi
  bridge: AiModelPickerBridge
  subscription: SubscriptionConnection
  lang: Lang
  unknownError: string
  enabled?: boolean
}) {
  const [settings, setSettings] = useState<AiSettings | null>(null)
  const [panel, setPanel] = useState<Connected | null>(null)
  const [failed, setFailed] = useState(false)
  const [attempt, setAttempt] = useState(0)
  const panelRef = useRef<Connected | null>(null)
  const live = useRef({ settings, lang, unknownError, refresh: subscription.refresh })
  live.current = { settings, lang, unknownError, refresh: subscription.refresh }
  const connectionRef = useRef<WorkspaceConnection | null>(null)
  if (!connectionRef.current)
    connectionRef.current = new WorkspaceConnection({
      settings: () => live.current.settings,
      lang: () => live.current.lang,
      unknownError: () => live.current.unknownError,
      confirm: (text) => window.confirm(text),
      onError: (error) => {
        if (shouldRecheckAfter(runErrorCauseOf(error, error.message)))
          live.current.refresh(error.provider ? [error.provider] : undefined)
      },
    })
  const connection = connectionRef.current
  const last = useRef(new Map<string, WorkspacePendingSend>())
  const resume = async (request: WorkspacePendingSend) => {
    const connected = panelRef.current
    if (!connected) return
    const run = await connected.host.resume(request)
    if (!run || connected !== panelRef.current) return
    const state = connected.host.conversations.state(request.conversationId)
    if (JSON.stringify(state.composer) === JSON.stringify(request.input))
      connected.host.conversations.edit(request.conversationId, {
        composer: { ...state.composer, text: '', attachments: [], skills: [] },
      })
  }
  const gate = useAiSendGate<WorkspacePendingSend>({
    settings,
    statuses: subscription.statuses,
    store: bridge,
    start: (request) => {
      void resume(request).catch(() => {})
    },
    onSettingsSaved: (next) => {
      live.current.settings = next
      setSettings(next)
    },
  })
  live.current.settings = gate.settings
  const gateRef = useRef(gate)
  gateRef.current = gate
  useSyncExternalStore(connection.subscribe, connection.version, connection.version)
  useSyncExternalStore(
    panel?.host.subscribe ?? idleSubscribe,
    panel?.host.version ?? idleVersion,
    idleVersion,
  )

  useEffect(() => {
    if (!enabled) return
    let alive = true
    let epoch = 0
    const refresh = async () => {
      const request = ++epoch
      try {
        const next = await bridge.getSettings()
        if (alive && epoch === request) setSettings(next)
      } catch {
        if (alive && epoch === request) setSettings(null)
      }
    }
    void refresh()
    const off = bridge.onSettingsChanged?.(() => {
      void refresh()
    })
    return () => {
      alive = false
      off?.()
    }
  }, [bridge, enabled])

  useEffect(() => {
    if (!enabled) return
    let alive = true
    let connected: Connected | null = null
    let off: (() => void) | undefined
    const requests = last.current
    const admittedRequests = new Map<string, WorkspacePendingSend>()
    setFailed(false)
    void connectWorkspacePanel(api, {
      lang: () => live.current.lang,
      gate: () => gateRef.current.peek(),
      admit: (request, queued) => {
        const admitted = queued
          ? gateRef.current.admitQueued(() => {})
          : gateRef.current.admit(request)
        if (queued && admitted)
          admittedRequests.set(request.conversationId, structuredClone(request))
        return admitted
      },
      toolsAllowed: connection.toolsAllowed,
      createTransport: (id, client, purpose) => connection.transport(id, client, purpose),
      prepare: async (_input, run) => {
        const id = run.binding.conversationId
        const request = admittedRequests.get(id)
        return {
          ...connection.prepare(run),
          onStart() {
            if (request) requests.set(id, request)
          },
          beforeTurn: async () => {
            if (gateRef.current.peek().kind !== 'run') throw new Error(live.current.unknownError)
          },
        }
      },
      onUsage: connection.usage,
      errorText: (error) =>
        error instanceof Error
          ? error.message
          : typeof error === 'string'
            ? error
            : live.current.unknownError,
    })
      .then((result) => {
        connected = result
        if (!alive) {
          void result.dispose().catch(() => {})
          return
        }
        panelRef.current = result
        off = result.client.subscribe((event) => {
          if (event.kind !== 'closed' || !alive) return
          panelRef.current = null
          gateRef.current.reset()
          requests.clear()
          setPanel(null)
          setFailed(true)
        })
        setPanel(result)
      })
      .catch(() => {
        if (alive) setFailed(true)
      })
    return () => {
      alive = false
      off?.()
      panelRef.current = null
      setPanel(null)
      gateRef.current.reset()
      requests.clear()
      void connected?.dispose().catch(() => {})
    }
  }, [api, connection, enabled, attempt])

  useEffect(() => {
    panel?.host.wake()
  }, [panel, gate.gate.kind, settings, subscription.statuses])

  const id = panel?.host.conversations.selectedId
  const feedback = id ? connection.state(id) : null
  const blocked = askBlockedRunOf(feedback?.error)
  return {
    panel,
    failed,
    reconnect: () => setAttempt((value) => value + 1),
    connection,
    composerControls: (
      <AiModelPicker
        bridge={bridge}
        lang={lang}
        subscriptionStatuses={subscription.statuses}
        onRefreshSubscriptions={() => subscription.refresh()}
      />
    ),
    notices: (
      <>
        <p className="ai-sub-note" role="note">
          {workspacePanelText(lang, 'sendScope')}
        </p>
        {gate.connectChoice && (
          <AiConnectPrompt
            choice={gate.connectChoice}
            statuses={subscription.statuses}
            lang={lang}
            disabled={subscription.busy}
            attention={gate.connectAttention}
            onOpenSettings={() => bridge.openModelSettings?.()}
            onInstall={subscription.install}
            onUpdate={subscription.update}
            onCancelLogin={subscription.cancelLogin}
            onLogin={subscription.login}
            onRecheck={subscription.recheck}
            onAcceptMetered={subscription.acceptMetered}
            onOpenInstallHelp={subscription.openInstallHelp}
            onOpenConnectHelp={subscription.openConnectHelp}
          />
        )}
        {gate.noticeDue && (
          <AiFirstConnectionNotice
            lang={lang}
            codexKeepsHistory
            onAcknowledge={gate.acknowledgeNotice}
          />
        )}
        {feedback?.truncated && (
          <AiInputTruncatedNotice
            lang={lang}
            onDismiss={() => {
              if (id) connection.dismissTruncated(id)
            }}
          />
        )}
        {gate.settings && localModelCapabilities(gate.settings).tools === false && (
          <AiToolsUnsupportedNotice lang={lang} />
        )}
        {blocked && (
          <AiBlockedNotice
            {...blocked}
            lang={lang}
            settings={gate.settings ?? {}}
            statuses={subscription.statuses}
            onContinueWith={(provider) => {
              const request = id ? last.current.get(id) : undefined
              if (request)
                gate.continueWith(provider, () => {
                  void resume(request).catch(() => {})
                })
            }}
          />
        )}
        {feedback && <AiSubscriptionUsageLine usage={feedback.subscriptionUsage} lang={lang} />}
        {feedback?.usage && <AiUsageChip usage={feedback.usage} session lang={lang} />}
      </>
    ),
  }
}
