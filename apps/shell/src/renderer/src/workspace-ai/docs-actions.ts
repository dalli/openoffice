import { createI18n, type Lang } from '@genoffice/i18n'
import { aiStrings } from '../../../../../docs/src/renderer/i18n/strings-ai'
import { workspacePanelText } from '../i18n/strings-workspace-panel'

const translate = createI18n(aiStrings)
export const workspaceDocText = translate

/** Reuse all 21 editor translations; these buttons fill a draft and never send it. */
export function workspaceDocsActions(lang: Lang) {
  return [
    { label: translate(lang, 'aiSummarizeBtn'), prompt: translate(lang, 'aiSummarizePrompt') },
    { label: translate(lang, 'aiPolishBtn'), prompt: workspacePanelText(lang, 'polishPrompt') },
    { label: translate(lang, 'aiTidyBtn'), prompt: translate(lang, 'aiTidyPrompt') },
    { label: translate(lang, 'aiMindMapBtn'), prompt: translate(lang, 'aiMindMapPrompt') },
  ]
}

export const workspaceMindMapQuestion = (lang: Lang, label: string) =>
  translate(lang, 'aiMindMapAskPrompt', { label })
