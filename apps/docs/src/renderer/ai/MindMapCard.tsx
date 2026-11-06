import { useI18n } from '../i18n/locale'
import { MindMapCardView, type MindMapCardProps } from './MindMapCardView'

/** The editor wrapper supplies its current locale; the shell uses the standalone view. */
export function MindMapCard(props: Omit<MindMapCardProps, 'lang'>) {
  const { lang } = useI18n()
  return <MindMapCardView {...props} lang={lang} />
}
