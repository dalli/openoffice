import { describe, expect, it, vi } from 'vitest'
import type { SkillHandoff } from '@genoffice/electron-utils/generate-handoff'
import { routeTemplateHandoff } from '../src/main/user-skills/template-handoff'

const handoff: SkillHandoff = {
  id: 'one',
  mode: 'skill',
  target: 'docs',
  text: 'Fill this copy',
  chips: [{ id: 'app:one', source: 'app', title: 'One' }],
  createdAt: 1,
}

describe('template handoff ownership', () => {
  it('consumes only an acknowledged Word transfer and leaves unavailable hosts to the editor', async () => {
    expect(await routeTemplateHandoff(handoff)).toBe(handoff)
    expect(await routeTemplateHandoff(handoff, async () => false)).toBe(handoff)
    expect(await routeTemplateHandoff(handoff, async () => true)).toBeNull()
    expect(
      await routeTemplateHandoff(handoff, async () => {
        throw new Error('Host unavailable')
      }),
    ).toBe(handoff)
    const route = vi.fn(async () => true)
    expect(await routeTemplateHandoff(null, route)).toBeNull()
    const slides = { ...handoff, target: 'slides' as const }
    expect(await routeTemplateHandoff(slides, route)).toBe(slides)
    const preview = { ...handoff, preview: { previewId: 'preview' } } as SkillHandoff
    expect(await routeTemplateHandoff(preview, route)).toBe(preview)
    expect(route).not.toHaveBeenCalled()
  })
})
