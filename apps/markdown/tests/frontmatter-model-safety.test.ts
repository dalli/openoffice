import { afterEach, describe, expect, it } from 'vitest'
import { Editor } from '@tiptap/core'
import StarterKit from '@tiptap/starter-kit'
import { executeTool } from '../src/renderer/ai/tools'

const editors: Editor[] = []
afterEach(() => {
  for (const editor of editors.splice(0)) editor.destroy()
})

describe('frontmatter model reads', () => {
  it.each([
    'private: <span data-redaction="" data-label="account">synthetic-secret-account</span>',
    'private: <span DATA-REDACTION="" data-label="account">synthetic-secret-account',
  ])('refuses raw private markup that the old tool returned verbatim', (yaml) => {
    const editor = new Editor({ extensions: [StarterKit] })
    editors.push(editor)
    const result = executeTool(
      editor,
      { id: 'read', name: 'read_frontmatter', input: {} },
      undefined,
      { read: () => yaml, write: () => {} },
    )
    expect(result).toMatchObject({ isError: true })
    expect((result as { output: string }).output).not.toContain('synthetic-secret-account')
    expect((result as { output: string }).output).not.toContain('data-label')
  })

  it('preserves unmarked YAML exactly without parsing or rewriting document metadata', () => {
    const editor = new Editor({ extensions: [StarterKit] })
    editors.push(editor)
    const yaml = '# user comment\ntitle: Draft\ntags: [one, two]\n'
    const result = executeTool(
      editor,
      { id: 'read', name: 'read_frontmatter', input: {} },
      undefined,
      { read: () => yaml, write: () => {} },
    )
    expect(result).toMatchObject({ output: yaml, mutated: false })
  })
})
