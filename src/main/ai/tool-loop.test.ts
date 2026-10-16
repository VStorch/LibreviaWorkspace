import { describe, expect, it } from 'vitest'
import type { AiPrompt, AiProvider, AiStreamPart } from '@services/ai/provider.js'
import { MAX_AI_TOOL_CALLS, MAX_AI_TOOL_ROUNDS } from '@shared/limits.js'
import { relayWithTools } from './tool-loop.js'

const PROMPT: AiPrompt = {
  model: 'm',
  system: 'sys',
  messages: [{ role: 'user', text: 'Revise.' }],
  maxTokens: 100,
}
const PARAGRAPHS = ['O gato dorme.', 'O cão late.']
const DOCUMENT = { paragraphs: PARAGRAPHS, readBudget: 100_000 }

const toolCall = (id: string, name: string, args: unknown): AiStreamPart => ({
  toolCall: { id, name, arguments: JSON.stringify(args) },
})

function scripted(rounds: AiStreamPart[][], seen: AiPrompt[] = []): AiProvider {
  return {
    listModels: async () => [],
    async *stream(prompt) {
      seen.push(prompt)
      yield* rounds[seen.length - 1] ?? [{ text: 'fim' }]
    },
  }
}

describe('laço de ferramentas', () => {
  it('roda as ferramentas, devolve os resultados na rodada seguinte e junta as propostas', async () => {
    const seen: AiPrompt[] = []
    const provider = scripted(
      [
        [
          { text: 'Vou ler.' },
          toolCall('a', 'search_document', { query: 'gato' }),
          toolCall('b', 'propose_edit', { paragraph: 1, find: 'dorme', replace: 'cochila' }),
        ],
        [toolCall('c', 'add_comment', { paragraph: 2, find: 'late', comment: 'Alto?' })],
        [{ text: 'Pronto.' }],
      ],
      seen,
    )
    const pushed: string[] = []
    const result = await relayWithTools(provider, PROMPT, DOCUMENT, new AbortController().signal, (text) =>
      pushed.push(text),
    )
    expect(seen).toHaveLength(3)
    expect(seen[0]?.tools?.map((tool) => tool.name)).toContain('propose_edit')
    expect(seen[0]?.exchanges).toEqual([])
    expect(seen[1]?.exchanges?.[0]).toMatchObject({
      text: 'Vou ler.',
      results: [
        { callId: 'a', content: '[1] O gato dorme.', isError: false },
        { callId: 'b', isError: false },
      ],
    })
    expect(seen[2]?.exchanges).toHaveLength(2)
    expect(pushed.join('')).toBe('Vou ler.\n\nPronto.')
    expect(result.proposals).toEqual([
      {
        id: '1',
        kind: 'edit',
        paragraph: 1,
        paragraphText: 'O gato dorme.',
        find: 'dorme',
        replace: 'cochila',
      },
      {
        id: '2',
        kind: 'comment',
        paragraph: 2,
        paragraphText: 'O cão late.',
        find: 'late',
        comment: 'Alto?',
      },
    ])
  })

  it('erro de ferramenta volta ao modelo, que pode tentar de novo', async () => {
    const seen: AiPrompt[] = []
    const provider = scripted(
      [
        [toolCall('a', 'propose_edit', { paragraph: 1, find: 'late', replace: 'x' })],
        [toolCall('b', 'propose_edit', { paragraph: 2, find: 'late', replace: 'ladra' })],
      ],
      seen,
    )
    const result = await relayWithTools(
      provider,
      PROMPT,
      DOCUMENT,
      new AbortController().signal,
      () => undefined,
    )
    expect(seen[1]?.exchanges?.[0]?.results[0]).toMatchObject({ callId: 'a', isError: true })
    expect(result.proposals).toHaveLength(1)
  })

  it(`para depois de ${MAX_AI_TOOL_ROUNDS} rodadas com o que tiver`, async () => {
    const seen: AiPrompt[] = []
    const always = Array.from({ length: 20 }, (_, round) => [
      toolCall(`c${round}`, 'add_comment', { paragraph: 1, find: 'gato', comment: `n${round}` }),
    ])
    const result = await relayWithTools(
      scripted(always, seen),
      PROMPT,
      DOCUMENT,
      new AbortController().signal,
      () => undefined,
    )
    expect(seen).toHaveLength(MAX_AI_TOOL_ROUNDS)
    expect(result.proposals).toHaveLength(MAX_AI_TOOL_ROUNDS)
  })

  it('cancelar no meio interrompe sem pedir outra rodada', async () => {
    const controller = new AbortController()
    let rounds = 0
    const provider: AiProvider = {
      listModels: async () => [],
      async *stream(_prompt, signal) {
        rounds++
        yield toolCall('a', 'search_document', { query: 'gato' })
        controller.abort()
        signal.throwIfAborted()
      },
    }
    await expect(
      relayWithTools(provider, PROMPT, DOCUMENT, controller.signal, () => undefined),
    ).rejects.toThrow()
    expect(rounds).toBe(1)
  })

  it('soma o custo de todas as rodadas', async () => {
    const provider = scripted([
      [toolCall('a', 'search_document', { query: 'gato' }), { costUsd: 0.25 }],
      [{ text: 'Pronto.' }, { costUsd: 0.5 }],
    ])
    const result = await relayWithTools(
      provider,
      PROMPT,
      DOCUMENT,
      new AbortController().signal,
      () => undefined,
    )
    expect(result.costUsd).toBe(0.75)
  })

  it(`responde com erro às chamadas além de ${MAX_AI_TOOL_CALLS} numa rodada`, async () => {
    const seen: AiPrompt[] = []
    const many = Array.from({ length: MAX_AI_TOOL_CALLS + 2 }, (_, index) =>
      toolCall(`c${index}`, 'search_document', { query: 'gato' }),
    )
    await relayWithTools(
      scripted([many], seen),
      PROMPT,
      DOCUMENT,
      new AbortController().signal,
      () => undefined,
    )
    const results = seen[1]?.exchanges?.[0]?.results ?? []
    expect(results).toHaveLength(MAX_AI_TOOL_CALLS + 2)
    expect(results.filter((result) => result.isError)).toHaveLength(2)
    expect(results[MAX_AI_TOOL_CALLS]?.callId).toBe(`c${MAX_AI_TOOL_CALLS}`)
  })

  it('acabado o orçamento de leitura, a ferramenta avisa em vez de ler', async () => {
    const seen: AiPrompt[] = []
    const provider = scripted(
      [
        [toolCall('a', 'read_paragraphs', { from: 1, count: 2 })],
        [toolCall('b', 'read_paragraphs', { from: 1, count: 2 })],
      ],
      seen,
    )
    const document = { paragraphs: PARAGRAPHS, readBudget: 40 }
    await relayWithTools(provider, PROMPT, document, new AbortController().signal, () => undefined)
    expect(seen[1]?.exchanges?.[0]?.results[0]).toMatchObject({ isError: false })
    expect(seen[2]?.exchanges?.[1]?.results[0]).toMatchObject({
      isError: true,
      content: expect.stringContaining('reading limit') as unknown,
    })
  })
})
