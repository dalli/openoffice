import { useMemo, useSyncExternalStore } from 'react'
import type { Lang } from '@genoffice/i18n'
import {
  SkillCardLine,
  SkillProposalCard,
  OriginalTemplateActions,
  SkillCardPreviewLine,
  skillCardBridge,
  skillTemplateText,
  readStoredSkillCards,
  type SkillCardBridge,
} from '@genoffice/ui'
import type { WorkspaceRunHost } from './run-host'
import type { WorkspaceSkillCard, WorkspaceUserSkills } from './user-skills'

const noSubscription = () => () => {}
const noVersion = () => 0
export function useWorkspaceSkills(skills: WorkspaceUserSkills | undefined) {
  useSyncExternalStore(
    skills?.subscribe ?? noSubscription,
    skills?.version ?? noVersion,
    skills?.version ?? noVersion,
  )
}

function WorkspaceSkillProposal({
  skills,
  card,
  id,
  lang,
  base,
  busy,
}: {
  skills: WorkspaceUserSkills
  card: WorkspaceSkillCard
  id: string
  lang: Lang
  base: SkillCardBridge | null
  busy: boolean
}) {
  const bridge = useMemo(
    () =>
      base &&
      ({
        ...base,
        saveFromCard: (input) => skills.save(id, card, base, input),
      } satisfies SkillCardBridge),
    [base, skills, id, card],
  )
  const document = useMemo(() => skills.document(card), [skills, card])
  return (
    <fieldset className="workspace-ai-skill-save" disabled={card.saving}>
      <SkillProposalCard
        lang={lang}
        proposal={card.proposal}
        bridge={bridge}
        canBeTemplate={!!document}
        initialDraft={card.draft}
        renderTemplateActions={
          document
            ? (context) => (
                <OriginalTemplateActions
                  context={{ ...context, busy: context.busy || busy }}
                  document={document}
                  bridge={skills.templateBridge}
                  onPreviewOpened={(opened) => skills.previewOpened(card, opened)}
                />
              )
            : undefined
        }
        onDraftChange={(draft) => skills.draft(card, draft)}
        onSaved={() => {}}
        onCancel={() => skills.cancel(card)}
      />
    </fieldset>
  )
}

/** Cards keep their draft in the conversation service even when the panel is folded or switched. */
export function WorkspaceSkillCards({
  skills,
  host,
  id,
  lang,
}: {
  skills: WorkspaceUserSkills
  host: WorkspaceRunHost
  id: string
  lang: Lang
}) {
  const cards = skills.proposals(id)
  const messages = host.conversations.state(id).messages
  const stored = readStoredSkillCards(messages, id).proposals
  const sequences = new Map<string, number>(
    messages.flatMap((message, i) =>
      (message.tools ?? []).map((_, j) => [`${id}-${i}-${j}`, message.seq] as const),
    ),
  )
  const text = (key: 'takenOver' | 'goToTab' | 'openWithTemplate') => skillTemplateText(lang, key)
  return (
    <div className="workspace-ai-skill-cards">
      {stored
        .filter(
          (record) =>
            !cards.some(
              (card) =>
                (card.proposal.name || card.proposal.title) === record.title &&
                sequences.get(record.key)! > card.beforeSeq &&
                sequences.get(record.key)! <= (card.afterSeq ?? -1),
            ),
        )
        .map((record) => (
          <SkillCardLine
            key={record.key}
            lang={lang}
            title={record.title}
            status={record.status}
            onOpenSettings={() => void skills.bridge.openSettings().catch(() => {})}
          />
        ))}
      {cards.map((card) =>
        card.status === 'open' ? (
          <WorkspaceSkillProposal
            key={card.proposal.key}
            skills={skills}
            id={id}
            card={card}
            lang={lang}
            base={skillCardBridge()}
            busy={host.busy()}
          />
        ) : card.status === 'preview' ? (
          <SkillCardPreviewLine
            key={card.proposal.key}
            lang={lang}
            title={card.draft.name || card.proposal.title}
            onGoToTab={() =>
              void skills.templateBridge?.focusTab(card.continuesIn!.tabId).catch(() => {})
            }
          />
        ) : (
          <SkillCardLine
            key={card.proposal.key}
            lang={lang}
            title={card.draft.name || card.proposal.title}
            status={card.status}
            onOpenSettings={() => void skills.bridge.openSettings().catch(() => {})}
          />
        ),
      )}
      {skills.templateNotes(id).map((note) => (
        <div key={note.key} className="ai-skill-card-line" dir="auto">
          <span>
            {note.tabId ? text('takenOver') : note.title}
            {note.several
              ? ` · ${skillTemplateText(lang, 'firstTemplateOnly', { title: note.title })}`
              : ''}
          </span>
          {' · '}
          <button
            type="button"
            className="ai-skill-card-link"
            disabled={note.opening}
            onClick={() => {
              if (note.tabId) void skills.templateBridge?.focusTab(note.tabId).catch(() => {})
              else void skills.openTemplate(id, note)
            }}
          >
            {text(note.tabId ? 'goToTab' : 'openWithTemplate')}
          </button>
        </div>
      ))}
    </div>
  )
}
