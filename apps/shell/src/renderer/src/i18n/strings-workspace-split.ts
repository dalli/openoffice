import { createI18n, defineStrings, type Lang, type Params } from '@genoffice/i18n'
import { zh } from './workspace-split/zh'
import { en } from './workspace-split/en'
import { ja } from './workspace-split/ja'
import { ko } from './workspace-split/ko'
import { fr } from './workspace-split/fr'
import { de } from './workspace-split/de'
import { es } from './workspace-split/es'
import { th } from './workspace-split/th'
import { id } from './workspace-split/id'
import { ru } from './workspace-split/ru'
import { ar } from './workspace-split/ar'
import { pt } from './workspace-split/pt'
import { it } from './workspace-split/it'
import { pl } from './workspace-split/pl'
import { cs } from './workspace-split/cs'
import { nl } from './workspace-split/nl'
import { ms } from './workspace-split/ms'
import { he } from './workspace-split/he'
import { hi } from './workspace-split/hi'
import { zhTW } from './workspace-split/zh-TW'
import { vi } from './workspace-split/vi'

/**
 * Window split controls shared by the tab strip and the native menu.
 * Sharded per language; `zh` defines the key set.
 */
export const workspaceSplitStrings = defineStrings({
  zh,
  en,
  ja,
  ko,
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

export type WorkspaceSplitStringKey = keyof typeof zh

const translate = createI18n(workspaceSplitStrings)

export function workspaceSplitText(
  lang: Lang,
  key: WorkspaceSplitStringKey,
  params?: Params,
): string {
  return translate(lang, key, params)
}
