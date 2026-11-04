import { useLayoutEffect, useRef, useState, useSyncExternalStore } from 'react'
import type { ReactNode } from 'react'
import { createPortal } from 'react-dom'
import { AiComposer, Markdown, aiSubscriptionText } from '@genoffice/ui'
import type { WorkspaceLayout } from '../../../shared/workspace-layout'
import { workspaceUiScale } from '../../../shared/workspace-layout'
import { workspaceConversationKey } from './conversations'
import type { WorkspaceConversationState } from './conversations'
import type { WorkspacePanelController, WorkspacePanelDocument } from './panel-controller'
import { WorkspaceRunCards } from './WorkspaceRunCards'
import { useI18n } from '../locale'
import { workspacePanelText } from '../i18n/strings-workspace-panel'
import { workspaceSplitText } from '../i18n/strings-workspace-split'
import './workspace-panel.css'

export interface WorkspacePanelMount {
  controller: WorkspacePanelController
  /** Main-acknowledged layout; the host reserves this rectangle before mounting shell UI. */
  layout: WorkspaceLayout
  uiScale: number
  setOpen(open: boolean): void
  /** Existing model/project/attachment/skill controls are supplied by their guarded service host. */
  composerControls?: ReactNode
  notices?: ReactNode
}

/** One shell surface. No default production host is installed until all service guards are wired. */
export function WorkspacePanel({ mount }: { mount: WorkspacePanelMount }) {
  const { controller, layout } = mount
  useSyncExternalStore(controller.subscribe, controller.version, controller.version)
  const { lang, t } = useI18n()
  const text = (key: Parameters<typeof workspacePanelText>[1]) => workspacePanelText(lang, key)
  const title = aiSubscriptionText(lang, 'aiAssistant')
  const documents = controller.documentList()
  const conversations = controller.host.conversations.list()
  const id = controller.host.conversations.selectedId
  const selected = conversations.find((item) => item.id === id)
  const selection = id ? controller.selection(id) : null
  const [referenceTab, setReferenceTab] = useState('')
  const [scope, setScope] = useState<'document' | 'selection'>('document')
  const scroll = useRef<HTMLDivElement>(null)

  // Inactive conversations keep their own position. Restoring never cancels their loop/cards.
  useLayoutEffect(() => {
    if (scroll.current && id)
      scroll.current.scrollTop = controller.host.conversations.state(id).scrollTop
    setReferenceTab('')
    setScope('document')
  }, [controller, id, layout.panelMode])

  const documentLabel = (doc: WorkspacePanelDocument): string => {
    const pane = doc.pane
      ? workspaceSplitText(
          lang,
          {
            left: 'splitLeftDocument',
            right: 'splitRightDocument',
            top: 'splitTopDocument',
            bottom: 'splitBottomDocument',
          }[doc.pane] as 'splitLeftDocument',
        )
      : ''
    // An explicit list position disambiguates even duplicate titles and copies of the same file.
    return `${doc.title} · ${doc.kind.toUpperCase()} · ${documents.indexOf(doc) + 1}${pane ? ` · ${pane}` : ''}${doc.available ? '' : ` · ${text('unavailable')}`}`
  }
  const stateLabel = (state: WorkspaceConversationState): string => {
    if (controller.host.permissions(state.id).requests().length) return text('permission')
    if (state.runs.some((run) => run.status === 'running')) return text('running')
    if (state.runs.length) return text('waiting')
    return ''
  }
  const conversationTitle = (state: WorkspaceConversationState): string =>
    controller.sourceFor(state.id)?.title || text('newConversation')
  const originLabel = (kind: WorkspaceConversationState['origin']['kind']): string =>
    text(
      {
        'legacy-document': 'originLegacy',
        project: 'originProject',
        workspace: 'originWorkspace',
        'window-memory': 'originMemory',
      }[kind] as 'originLegacy',
    )
  const options = controller.sourceList().map((source) => ({
    key: workspaceConversationKey(source.origin),
    source,
    state: conversations.find(
      (state) => workspaceConversationKey(state.origin) === workspaceConversationKey(source.origin),
    ),
  }))
  for (const state of conversations) {
    if (!options.some((item) => item.state?.id === state.id))
      options.push({
        key: workspaceConversationKey(state.origin),
        source: { origin: state.origin, title: conversationTitle(state) },
        state,
      })
  }
  const stop = (state: WorkspaceConversationState): void => {
    for (const run of state.runs) void controller.host.cancel(run.run)
  }
  const work = conversations.filter((state) => state.runs.length || stateLabel(state))
  const targetTab = selection?.target?.binding.document.tabId
  const targetClosed =
    !!targetTab && !documents.some((doc) => doc.id === targetTab && doc.available)
  const referencesClosed = selection?.references.some(
    (ref) => !documents.some((doc) => doc.id === ref.binding.document.tabId && doc.available),
  )
  const scale = workspaceUiScale(mount.uiScale)
  const bounds = layout.panel
  if (!bounds || layout.panelMode === 'hidden') return null

  return createPortal(
    <aside
      className="workspace-ai-surface"
      aria-label={title}
      style={{ left: bounds.x, top: bounds.y, width: bounds.width, height: bounds.height }}
    >
      <div
        className="workspace-ai-panel"
        style={{
          zoom: scale,
          height: bounds.height / scale,
          width: Math.max(0, bounds.width - 2) / scale,
        }}
      >
        {layout.panelMode === 'rail' ? (
          <>
            <button
              type="button"
              aria-label={text('open')}
              title={text('open')}
              onClick={() => mount.setOpen(true)}
            >
              AI
            </button>
            {work.map((state) => (
              <div key={state.id} className="workspace-ai-rail-work">
                <button
                  type="button"
                  title={`${conversationTitle(state)} · ${stateLabel(state)}`}
                  aria-label={`${conversationTitle(state)} · ${stateLabel(state)}`}
                  onClick={() => {
                    controller.select(state.id)
                    mount.setOpen(true)
                  }}
                >
                  •
                </button>
                <button
                  type="button"
                  aria-label={`${t('askStop')} · ${conversationTitle(state)}`}
                  onClick={() => stop(state)}
                >
                  ■
                </button>
              </div>
            ))}
          </>
        ) : (
          <>
            <header className="workspace-ai-header">
              <strong>{title}</strong>
              <button
                type="button"
                aria-label={text('collapse')}
                title={text('collapse')}
                onClick={() => mount.setOpen(false)}
              >
                ›
              </button>
            </header>
            <div className="workspace-ai-row">
              <label className="workspace-ai-grow">
                <span>{text('conversation')}</span>
                <select
                  value={selected ? workspaceConversationKey(selected.origin) : ''}
                  onChange={(event) => {
                    const option = options.find((item) => item.key === event.target.value)
                    if (option?.state) controller.select(option.state.id)
                    else if (option) void controller.open(option.source)
                  }}
                >
                  <option value="" disabled>
                    {text('none')}
                  </option>
                  {options.map(({ key, source, state }, index) => (
                    <option key={key} value={key}>
                      {source.title || text('newConversation')} · {originLabel(source.origin.kind)}
                      {source.projectTitle ? ` · ${source.projectTitle}` : ''} · {index + 1}
                      {state
                        ? ` · ${state.composer.target ? (documents.find((doc) => doc.id === state.composer.target!.document.tabId)?.title ?? text('closed')) : text('none')}${stateLabel(state) ? ` · ${stateLabel(state)}` : ''}${state.messages.at(-1)?.ts ? ` · ${new Date(state.messages.at(-1)!.ts).toLocaleString(lang)}` : ''}`
                        : ''}
                    </option>
                  ))}
                </select>
              </label>
              <button
                type="button"
                aria-label={text('newConversation')}
                title={text('newConversation')}
                onClick={() => {
                  void controller.create()
                }}
              >
                +
              </button>
            </div>
            {controller.hasOpenError() && <p role="alert">{text('selectionFailed')}</p>}
            {selected && id && selection ? (
              <>
                <label className="workspace-ai-field">
                  <span>{text('target')}</span>
                  <select
                    value={targetTab ?? ''}
                    aria-busy={!!selection.pending}
                    onChange={(event) => {
                      if (event.target.value) void controller.target(id, event.target.value)
                    }}
                  >
                    <option value="" disabled>
                      {text('none')}
                    </option>
                    {targetClosed && <option value={targetTab}>{text('closed')}</option>}
                    {documents.map((doc) => (
                      <option key={doc.id} value={doc.id} disabled={!doc.available}>
                        {documentLabel(doc)}
                      </option>
                    ))}
                  </select>
                </label>
                <fieldset className="workspace-ai-references">
                  <legend>{text('references')}</legend>
                  {selection.references.map((ref) => (
                    <div key={ref.id} className="workspace-ai-row">
                      <span className="workspace-ai-grow" dir="auto">
                        {ref.title} · {text(ref.scope)}
                        {documents.some(
                          (doc) => doc.id === ref.binding.document.tabId && doc.available,
                        )
                          ? ''
                          : ` · ${text('closed')}`}
                      </span>
                      <button
                        type="button"
                        aria-label={`${text('removeReference')} · ${ref.title}`}
                        onClick={() => controller.removeReference(id, ref.id)}
                      >
                        ×
                      </button>
                    </div>
                  ))}
                  <div className="workspace-ai-row">
                    <select
                      className="workspace-ai-grow"
                      aria-label={text('referenceDocument')}
                      value={referenceTab}
                      onChange={(event) => setReferenceTab(event.target.value)}
                    >
                      <option value="">{text('none')}</option>
                      {documents
                        .filter((doc) => doc.id !== targetTab)
                        .map((doc) => (
                          <option key={doc.id} value={doc.id} disabled={!doc.available}>
                            {documentLabel(doc)}
                          </option>
                        ))}
                    </select>
                    <select
                      aria-label={text('scope')}
                      value={scope}
                      onChange={(event) => setScope(event.target.value as 'document' | 'selection')}
                    >
                      <option value="document">{text('document')}</option>
                      <option value="selection">{text('selection')}</option>
                    </select>
                    <button
                      type="button"
                      aria-label={text('addReference')}
                      title={text('addReference')}
                      disabled={!referenceTab || !!selection.pending}
                      onClick={() => {
                        void controller.addReference(id, referenceTab, scope)
                      }}
                    >
                      +
                    </button>
                  </div>
                </fieldset>
                {(targetClosed || referencesClosed) && <p role="status">{text('closed')}</p>}
                {selection.error && <p role="alert">{text('selectionFailed')}</p>}
                <div
                  className="workspace-ai-transcript"
                  ref={scroll}
                  onScroll={(event) =>
                    controller.host.conversations.edit(id, {
                      scrollTop: event.currentTarget.scrollTop,
                    })
                  }
                >
                  {selected.messages.map((message, index) => (
                    <article
                      key={`${message.seq}:${index}`}
                      className={`workspace-ai-message workspace-ai-message-${message.role}`}
                      dir="auto"
                    >
                      <Markdown text={message.text} />
                    </article>
                  ))}
                  {selected.progress && (
                    <div role="status" dir="auto">
                      <Markdown text={selected.progress} />
                    </div>
                  )}
                  {selected.error && (
                    <p role="alert" dir="auto">
                      {selected.error}
                    </p>
                  )}
                  <WorkspaceRunCards
                    host={controller.host}
                    conversationId={id}
                    lang={lang}
                    labels={{
                      undo: text('undo'),
                      restoring: text('restoring'),
                      reverted: text('reverted'),
                      refused: text('refused'),
                    }}
                  />
                </div>
                {mount.notices}
                <AiComposer
                  value={selected.composer.text}
                  busy={selected.runs.length > 0}
                  sendDisabled={!controller.canSend(id)}
                  placeholder={text('placeholder')}
                  ariaLabel={text('placeholder')}
                  hintIdle=""
                  hintBusy=""
                  sendLabel={t('askSend')}
                  stopLabel={t('askStop')}
                  footerStart={mount.composerControls}
                  onChange={(value) =>
                    controller.host.conversations.edit(id, {
                      composer: {
                        ...controller.host.conversations.state(id).composer,
                        text: value,
                      },
                    })
                  }
                  onSend={() => {
                    void controller.send(id)
                  }}
                  onQueue={() => {
                    void controller.send(id)
                  }}
                  onStop={() => stop(selected)}
                />
              </>
            ) : (
              <p className="workspace-ai-empty">{text('empty')}</p>
            )}
            {work.length > 0 && (
              <nav className="workspace-ai-work" aria-label={text('running')}>
                {work.map((state) => (
                  <div className="workspace-ai-row" key={state.id}>
                    <button
                      type="button"
                      className="workspace-ai-grow"
                      onClick={() => controller.select(state.id)}
                    >
                      {conversationTitle(state)} · {stateLabel(state)}
                    </button>
                    <button
                      type="button"
                      aria-label={`${t('askStop')} · ${conversationTitle(state)}`}
                      onClick={() => stop(state)}
                    >
                      ■
                    </button>
                  </div>
                ))}
              </nav>
            )}
          </>
        )}
      </div>
    </aside>,
    document.body,
  )
}
