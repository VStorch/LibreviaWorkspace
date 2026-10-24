import { describe, expect, it } from 'vitest'
import { MAX_AI_PROPOSALS } from '@shared/limits.js'
import {
  AiToolName,
  APPLIED_DOCUMENT_TOOLS,
  appliedOutcome,
  DOCUMENT_TOOLS,
  runDocumentTool,
  type ToolContext,
} from './document-tools.js'

const CONTEXT: ToolContext = {
  paragraphs: ['O gato dorme.', '', 'O cão late e o cão corre.', 'Fim do Gato.'],
  proposalCount: 0,
}

const call = (name: string, args: unknown, context: ToolContext = CONTEXT) =>
  runDocumentTool(
    { id: 'c1', name, arguments: typeof args === 'string' ? args : JSON.stringify(args) },
    context,
  )

describe('document tools', () => {
  it('describe the five tools with an object JSON schema', () => {
    expect(DOCUMENT_TOOLS.map((tool) => tool.name)).toEqual([
      'read_paragraphs',
      'search_document',
      'propose_edit',
      'insert_content',
      'add_comment',
    ])
    for (const tool of DOCUMENT_TOOLS) {
      expect(tool.parameters).toMatchObject({ type: 'object' })
      expect(tool.parameters).not.toHaveProperty('$schema')
    }
  })

  it('read_paragraphs returns paragraphs numbered from 1', () => {
    expect(call(AiToolName.ReadParagraphs, { from: 3, count: 5 })).toEqual({
      content: '[3] O cão late e o cão corre.\n[4] Fim do Gato.',
      isError: false,
    })
    expect(call(AiToolName.ReadParagraphs, { from: 9, count: 1 }).isError).toBe(true)
    expect(call(AiToolName.ReadParagraphs, { from: 1, count: 51 }).isError).toBe(true)
  })

  it('search_document matches case-insensitively and respects the maximum', () => {
    expect(call(AiToolName.SearchDocument, { query: 'gato' }).content).toBe(
      '[1] O gato dorme.\n[4] Fim do Gato.',
    )
    expect(call(AiToolName.SearchDocument, { query: 'GATO', max: 1 }).content).toBe('[1] O gato dorme.')
    expect(call(AiToolName.SearchDocument, { query: 'zebra' })).toEqual({
      content: 'No paragraph contains that text.',
      isError: false,
    })
    expect(call(AiToolName.SearchDocument, { query: 'x'.repeat(201) }).isError).toBe(true)
  })

  it('a valid propose_edit becomes a proposal with the reason', () => {
    const outcome = call(AiToolName.ProposeEdit, {
      paragraph: 1,
      paragraphText: 'O gato dorme.',
      find: 'dorme',
      replace: 'cochila',
      reason: 'tom',
    })
    expect(outcome.isError).toBe(false)
    expect(outcome.proposal).toEqual({
      id: '1',
      kind: 'edit',
      paragraph: 1,
      paragraphText: 'O gato dorme.',
      find: 'dorme',
      replace: 'cochila',
      reason: 'tom',
    })
  })

  it('propose_edit refuses a missing or repeated snippet, a nonexistent paragraph and an identical replacement', () => {
    const ausente = call(AiToolName.ProposeEdit, { paragraph: 1, find: 'late', replace: 'x' })
    expect(ausente).toMatchObject({ isError: true })
    expect(ausente.content).toContain('not in paragraph 1')
    expect(ausente.proposal).toBeUndefined()
    const repeated = call(AiToolName.ProposeEdit, { paragraph: 3, find: 'cão', replace: 'x' })
    expect(repeated.content).toContain('appears 2 times')
    expect(call(AiToolName.ProposeEdit, { paragraph: 7, find: 'a', replace: 'b' }).content).toContain(
      'no paragraph 7',
    )
    expect(call(AiToolName.ProposeEdit, { paragraph: 1, find: 'gato', replace: 'gato' }).isError).toBe(true)
  })

  it('validates the input limits', () => {
    expect(call(AiToolName.ProposeEdit, { paragraph: 1, find: 'x'.repeat(2001), replace: '' }).isError).toBe(
      true,
    )
    expect(
      call(AiToolName.ProposeEdit, { paragraph: 1, find: 'gato', replace: 'x'.repeat(10_001) }).isError,
    ).toBe(true)
    expect(
      call(AiToolName.ProposeEdit, { paragraph: 1, find: 'gato', replace: 'g', reason: 'x'.repeat(501) })
        .isError,
    ).toBe(true)
    expect(
      call(AiToolName.AddComment, { paragraph: 1, find: 'gato', comment: 'x'.repeat(2001) }).isError,
    ).toBe(true)
    expect(call(AiToolName.AddComment, { paragraph: 0, find: 'gato', comment: 'c' }).isError).toBe(true)
  })

  it('a valid add_comment becomes a comment proposal with the next number', () => {
    const outcome = call(
      AiToolName.AddComment,
      { paragraph: 4, find: 'Gato', comment: 'Maiúscula?' },
      { ...CONTEXT, proposalCount: 2 },
    )
    expect(outcome.proposal).toEqual({
      id: '3',
      kind: 'comment',
      paragraph: 4,
      paragraphText: 'Fim do Gato.',
      find: 'Gato',
      comment: 'Maiúscula?',
    })
  })

  it('past the proposal limit, returns an error and does not record', () => {
    const outcome = call(
      AiToolName.AddComment,
      { paragraph: 1, find: 'gato', comment: 'c' },
      { ...CONTEXT, proposalCount: MAX_AI_PROPOSALS },
    )
    expect(outcome.isError).toBe(true)
    expect(outcome.proposal).toBeUndefined()
  })

  it('unknown tool and invalid JSON are tool errors', () => {
    expect(call('delete_file', {})).toMatchObject({ isError: true })
    expect(call('delete_file', {}).content).toContain('read_paragraphs')
    expect(call(AiToolName.ReadParagraphs, '{"from": 1,')).toEqual({
      content: 'The arguments are not valid JSON.',
      isError: true,
    })
  })
})

describe('insert_content', () => {
  it('after a paragraph, with its text; after 0, at the start with no text', () => {
    expect(call(AiToolName.InsertContent, { after: 3, content: '## Novo\n\n- um' })).toEqual({
      content: 'Insertion 1 proposed; the person will decide.',
      isError: false,
      proposal: {
        id: '1',
        kind: 'insert',
        paragraph: 3,
        paragraphText: 'O cão late e o cão corre.',
        content: '## Novo\n\n- um',
      },
    })
    expect(call(AiToolName.InsertContent, { after: 0, content: 'Início.' }).proposal).toMatchObject({
      paragraph: 0,
      paragraphText: '',
    })
  })

  it('writes into an empty document', () => {
    const empty = { paragraphs: [''], proposalCount: 0 }
    expect(call(AiToolName.InsertContent, { after: 0, content: '# Título' }, empty).isError).toBe(false)
  })

  it('refuses a paragraph past the end, empty content and the proposal limit', () => {
    expect(call(AiToolName.InsertContent, { after: 5, content: 'x' }).content).toBe(
      'There is no paragraph 5; the document has 4.',
    )
    expect(call(AiToolName.InsertContent, { after: 1, content: '' }).isError).toBe(true)
    const full = { ...CONTEXT, proposalCount: MAX_AI_PROPOSALS }
    expect(call(AiToolName.InsertContent, { after: 1, content: 'x' }, full).content).toContain('limit')
  })

  it('tells the model the content is Markdown that becomes formatting', () => {
    const tool = DOCUMENT_TOOLS.find((candidate) => candidate.name === AiToolName.InsertContent)
    expect(tool?.description).toContain('empty document')
    expect(JSON.stringify(tool?.parameters)).toContain('Markdown')
  })
})

describe('tools for clients outside the app', () => {
  const edit = {
    id: '1',
    kind: 'edit' as const,
    paragraph: 2,
    paragraphText: 'O gato dorme.',
    find: 'gato',
    replace: 'cão',
  }

  it('the same tools, saying the change enters as a revision the person accepts or rejects', () => {
    expect(APPLIED_DOCUMENT_TOOLS.map((tool) => tool.name)).toEqual(DOCUMENT_TOOLS.map((tool) => tool.name))
    const proposeEdit = APPLIED_DOCUMENT_TOOLS.find((tool) => tool.name === AiToolName.ProposeEdit)
    expect(proposeEdit?.description).toContain('tracked revision')
    expect(proposeEdit?.description).toContain('accepts or rejects')
  })

  it('the result says whether the change went in or the paragraph changed', () => {
    expect(appliedOutcome(edit, true)).toMatchObject({ isError: false })
    expect(appliedOutcome(edit, false)).toEqual({
      content: 'Paragraph 2 changed before this went in: read it again and retry.',
      isError: true,
    })
  })
})
