import type { SkillHandoff } from '@genoffice/electron-utils/generate-handoff'

/** Only an acknowledged transfer consumes the editor's existing one-use request. */
export async function routeTemplateHandoff(
  handoff: SkillHandoff | null,
  route?: (handoff: SkillHandoff) => Promise<boolean>,
): Promise<SkillHandoff | null> {
  if (!handoff || handoff.target !== 'docs' || handoff.preview || !route) return handoff
  // The route must settle every emitted offer before refusing or throwing.
  try {
    return (await route(handoff)) ? null : handoff
  } catch {
    return handoff
  }
}
