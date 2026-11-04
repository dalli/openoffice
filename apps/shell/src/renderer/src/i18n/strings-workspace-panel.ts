import { createI18n, defineStrings, type Lang, type Params } from '@genoffice/i18n'
import { zh } from './workspace-panel/zh'
import { en } from './workspace-panel/en'
import { ko } from './workspace-panel/ko'
import { ja } from './workspace-panel/ja'
import { fr } from './workspace-panel/fr'
import { de } from './workspace-panel/de'
import { es } from './workspace-panel/es'
import { th } from './workspace-panel/th'
import { id } from './workspace-panel/id'
import { ru } from './workspace-panel/ru'
import { ar } from './workspace-panel/ar'
import { pt } from './workspace-panel/pt'
import { it } from './workspace-panel/it'
import { pl } from './workspace-panel/pl'
import { cs } from './workspace-panel/cs'
import { nl } from './workspace-panel/nl'
import { ms } from './workspace-panel/ms'
import { he } from './workspace-panel/he'
import { hi } from './workspace-panel/hi'
import { zhTW } from './workspace-panel/zh-TW'
import { vi } from './workspace-panel/vi'

export const workspacePanelStrings = defineStrings({
  zh,
  en,
  ko,
  ja,
  fr,
  de,
  es,
  th,
  id,
  ru,
  ar,
  pt,
  it,
  pl,
  cs,
  nl,
  ms,
  he,
  hi,
  'zh-TW': zhTW,
  vi,
})
export type WorkspacePanelStringKey = keyof typeof zh
const translate = createI18n(workspacePanelStrings)
export const workspacePanelText = (
  lang: Lang,
  key: WorkspacePanelStringKey,
  params?: Params,
): string => translate(lang, key, params)
