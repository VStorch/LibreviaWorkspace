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

describe('ferramentas do documento', () => {
  it('descrevem as quatro ferramentas com esquema JSON de objeto', () => {
    expect(DOCUMENT_TOOLS.map((tool) => tool.name)).toEqual([
      'read_paragraphs',
      'search_document',
      'propose_edit',
      'add_comment',
    ])
    for (const tool of DOCUMENT_TOOLS) {
      expect(tool.parameters).toMatchObject({ type: 'object' })
      expect(tool.parameters).not.toHaveProperty('$schema')
    }
  })

  it('read_paragraphs devolve os parágrafos numerados a partir de 1', () => {
    expect(call(AiToolName.ReadParagraphs, { from: 3, count: 5 })).toEqual({
      content: '[3] O cão late e o cão corre.\n[4] Fim do Gato.',
      isError: false,
    })
    expect(call(AiToolName.ReadParagraphs, { from: 9, count: 1 }).isError).toBe(true)
    expect(call(AiToolName.ReadParagraphs, { from: 1, count: 51 }).isError).toBe(true)
  })

  it('search_document acha sem diferenciar maiúsculas e respeita o máximo', () => {
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

  it('propose_edit válida vira proposta com o motivo', () => {
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

  it('propose_edit recusa trecho ausente, repetido, parágrafo inexistente e troca igual', () => {
    const ausente = call(AiToolName.ProposeEdit, { paragraph: 1, find: 'late', replace: 'x' })
    expect(ausente).toMatchObject({ isError: true })
    expect(ausente.content).toContain('not in paragraph 1')
    expect(ausente.proposal).toBeUndefined()
    const repetido = call(AiToolName.ProposeEdit, { paragraph: 3, find: 'cão', replace: 'x' })
    expect(repetido.content).toContain('appears 2 times')
    expect(call(AiToolName.ProposeEdit, { paragraph: 7, find: 'a', replace: 'b' }).content).toContain(
      'no paragraph 7',
    )
    expect(call(AiToolName.ProposeEdit, { paragraph: 1, find: 'gato', replace: 'gato' }).isError).toBe(true)
  })

  it('valida os limites das entradas', () => {
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

  it('add_comment válida vira proposta de comentário com o número seguinte', () => {
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

  it('passado o limite de propostas, devolve erro e não registra', () => {
    const outcome = call(
      AiToolName.AddComment,
      { paragraph: 1, find: 'gato', comment: 'c' },
      { ...CONTEXT, proposalCount: MAX_AI_PROPOSALS },
    )
    expect(outcome.isError).toBe(true)
    expect(outcome.proposal).toBeUndefined()
  })

  it('ferramenta desconhecida e JSON inválido são erros de ferramenta', () => {
    expect(call('delete_file', {})).toMatchObject({ isError: true })
    expect(call('delete_file', {}).content).toContain('read_paragraphs')
    expect(call(AiToolName.ReadParagraphs, '{"from": 1,')).toEqual({
      content: 'The arguments are not valid JSON.',
      isError: true,
    })
  })
})

describe('ferramentas para clientes de fora do aplicativo', () => {
  const edit = {
    id: '1',
    kind: 'edit' as const,
    paragraph: 2,
    paragraphText: 'O gato dorme.',
    find: 'gato',
    replace: 'cão',
  }

  it('as mesmas ferramentas, dizendo que a mudança entra como revisão que a pessoa aceita ou rejeita', () => {
    expect(APPLIED_DOCUMENT_TOOLS.map((tool) => tool.name)).toEqual(DOCUMENT_TOOLS.map((tool) => tool.name))
    const proposeEdit = APPLIED_DOCUMENT_TOOLS.find((tool) => tool.name === AiToolName.ProposeEdit)
    expect(proposeEdit?.description).toContain('tracked revision')
    expect(proposeEdit?.description).toContain('accepts or rejects')
  })

  it('o resultado diz se a mudança entrou ou se o parágrafo mudou', () => {
    expect(appliedOutcome(edit, true)).toMatchObject({ isError: false })
    expect(appliedOutcome(edit, false)).toEqual({
      content: 'Paragraph 2 changed before this went in: read it again and retry.',
      isError: true,
    })
  })
})
