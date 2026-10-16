import { describe, expect, it } from 'vitest'
import Anthropic from '@anthropic-ai/sdk'
import { ErrorCode } from '@shared/errors.js'
import { MAX_AI_TOOL_ARGUMENTS_LENGTH, MAX_AI_TOOL_CALLS } from '@shared/limits.js'
import { t } from '../i18n.js'
import { createAnthropicProvider } from './anthropic.js'

const PROMPT = {
  model: 'modelo-x',
  system: 'sys',
  messages: [{ role: 'user' as const, text: 'Olá' }],
  maxTokens: 100,
}

type Client = Parameters<typeof createAnthropicProvider>[1]

function clientWith(events: unknown[], calls: unknown[] = []): Client {
  return {
    beta: {
      messages: {
        stream: (body: unknown, options: unknown) => {
          calls.push(body, options)
          return (async function* () {
            for (const event of events) yield event
          })()
        },
      },
      models: {
        list: (params: unknown) => {
          calls.push(params)
          return (async function* () {
            yield { id: 'modelo-a', display_name: 'Modelo A', allowed_fallback_models: ['modelo-b'] }
            yield { id: 'modelo-b', display_name: 'Modelo B', allowed_fallback_models: [] }
          })()
        },
      },
    },
  } as unknown as Client
}

const FALLBACK = {
  type: 'content_block_start',
  index: 0,
  content_block: { type: 'fallback', from: { model: 'modelo-x' }, to: { model: 'modelo-y' } },
}

const text = (value: string): unknown => ({
  type: 'content_block_delta',
  delta: { type: 'text_delta', text: value },
})

async function collect(iterable: AsyncIterable<unknown>): Promise<unknown[]> {
  const parts: unknown[] = []
  for await (const part of iterable) parts.push(part)
  return parts
}

const CONVERSATION = {
  ...PROMPT,
  messages: [
    { role: 'user' as const, text: 'Qual o tema?' },
    { role: 'assistant' as const, text: 'Chuva.' },
    { role: 'user' as const, text: 'E o tom?' },
  ],
}
describe('provedor Anthropic', () => {
  it('manda a conversa inteira como mensagens, com o system à parte', async () => {
    const calls: unknown[] = []
    const provider = createAnthropicProvider('chave', clientWith([], calls))
    await collect(provider.stream(CONVERSATION, new AbortController().signal))
    expect(calls[0]).toMatchObject({
      system: 'sys',
      messages: [
        { role: 'user', content: 'Qual o tema?' },
        { role: 'assistant', content: 'Chuva.' },
        { role: 'user', content: 'E o tom?' },
      ],
    })
  })

  it('lê só os pedaços de texto e passa o sinal de cancelamento', async () => {
    const calls: unknown[] = []
    const signal = new AbortController().signal
    const provider = createAnthropicProvider(
      'chave',
      clientWith(
        [
          { type: 'message_start' },
          { type: 'content_block_delta', delta: { type: 'text_delta', text: 'Olá' } },
          { type: 'content_block_delta', delta: { type: 'input_json_delta', partial_json: '{' } },
          { type: 'message_delta', delta: { stop_reason: 'end_turn' } },
        ],
        calls,
      ),
    )

    expect(await collect(provider.stream(PROMPT, signal))).toEqual([{ text: 'Olá' }])
    expect(calls[0]).toMatchObject({ model: 'modelo-x', max_tokens: 100, system: 'sys' })
    expect(calls[1]).toEqual({ signal })
    expect(calls[0]).not.toHaveProperty('fallbacks')
  })

  it('com o recurso ligado, pede a troca de modelo e passa adiante o modelo que assumiu', async () => {
    const calls: unknown[] = []
    const provider = createAnthropicProvider('chave', clientWith([text('Come'), FALLBACK, text('ço')], calls))
    const parts = await collect(provider.stream({ ...PROMPT, fallback: true }, new AbortController().signal))
    expect(parts).toEqual([{ text: 'Come' }, { model: 'modelo-y' }, { text: 'ço' }])
    expect(calls[0]).toMatchObject({ betas: ['server-side-fallback-2026-07-01'], fallbacks: 'default' })
  })

  it('se o modelo que assumiu também recusar, o erro continua legível', async () => {
    const provider = createAnthropicProvider(
      'chave',
      clientWith([FALLBACK, { type: 'message_delta', delta: { stop_reason: 'refusal' } }]),
    )
    await expect(
      collect(provider.stream({ ...PROMPT, fallback: true }, new AbortController().signal)),
    ).rejects.toMatchObject({ code: ErrorCode.AiRefused })
  })

  it('recusa do modelo vira erro legível', async () => {
    const provider = createAnthropicProvider(
      'chave',
      clientWith([{ type: 'message_delta', delta: { stop_reason: 'refusal' } }]),
    )
    await expect(collect(provider.stream(PROMPT, new AbortController().signal))).rejects.toMatchObject({
      code: ErrorCode.AiRefused,
    })
  })

  it('erro de autenticação do SDK vira erro de chave', async () => {
    const client = {
      beta: {
        messages: {
          stream: () => {
            throw new Anthropic.AuthenticationError(401, undefined, 'invalid x-api-key', new Headers())
          },
        },
      },
    } as unknown as Client
    await expect(
      collect(createAnthropicProvider('chave', client).stream(PROMPT, new AbortController().signal)),
    ).rejects.toMatchObject({ code: ErrorCode.AiAuthFailed })
  })

  it('lista os modelos com o nome de exibição e marca os que aceitam outro modelo', async () => {
    const calls: unknown[] = []
    const provider = createAnthropicProvider('chave', clientWith([], calls))
    expect(await provider.listModels(new AbortController().signal)).toEqual([
      { id: 'modelo-a', label: 'Modelo A', fallback: true },
      { id: 'modelo-b', label: 'Modelo B', fallback: false },
    ])
    expect(calls[0]).toEqual({ betas: ['server-side-fallback-2026-06-01'] })
  })

  it('monta a chamada de ferramenta a partir de tool_use e input_json_delta', async () => {
    const provider = createAnthropicProvider(
      'k',
      clientWith([
        text('Vou ver.'),
        {
          type: 'content_block_start',
          index: 1,
          content_block: { type: 'tool_use', id: 't1', name: 'search_document', input: {} },
        },
        {
          type: 'content_block_delta',
          index: 1,
          delta: { type: 'input_json_delta', partial_json: '{"query":' },
        },
        {
          type: 'content_block_delta',
          index: 1,
          delta: { type: 'input_json_delta', partial_json: '"gato"}' },
        },
        { type: 'content_block_stop', index: 1 },
      ]),
    )
    expect(await collect(provider.stream(PROMPT, new AbortController().signal))).toEqual([
      { text: 'Vou ver.' },
      { toolCall: { id: 't1', name: 'search_document', arguments: '{"query":"gato"}' } },
    ])
  })

  it('aborta a resposta que pede ferramentas demais', async () => {
    const events = Array.from({ length: MAX_AI_TOOL_CALLS + 1 }, (_, index) => [
      {
        type: 'content_block_start',
        index,
        content_block: { type: 'tool_use', id: `t${index}`, name: 'search_document', input: {} },
      },
      { type: 'content_block_stop', index },
    ]).flat()
    const provider = createAnthropicProvider('k', clientWith(events))
    await expect(collect(provider.stream(PROMPT, new AbortController().signal))).rejects.toMatchObject({
      message: t('errors.ai.toolCallsTooLarge'),
    })
  })

  it('aborta argumentos de ferramenta longos demais', async () => {
    const provider = createAnthropicProvider(
      'k',
      clientWith([
        {
          type: 'content_block_start',
          index: 0,
          content_block: { type: 'tool_use', id: 't1', name: 'search_document', input: {} },
        },
        {
          type: 'content_block_delta',
          index: 0,
          delta: { type: 'input_json_delta', partial_json: 'x'.repeat(MAX_AI_TOOL_ARGUMENTS_LENGTH + 1) },
        },
      ]),
    )
    await expect(collect(provider.stream(PROMPT, new AbortController().signal))).rejects.toMatchObject({
      message: t('errors.ai.toolCallsTooLarge'),
    })
  })

  it('manda as ferramentas e o histórico como tool_use e tool_result', async () => {
    const calls: unknown[] = []
    const provider = createAnthropicProvider('k', clientWith([], calls))
    const prompt = {
      ...PROMPT,
      tools: [
        { name: 'search_document', description: 'Busca.', parameters: { type: 'object', properties: {} } },
      ],
      exchanges: [
        {
          text: 'Vou ver.',
          calls: [{ id: 't1', name: 'search_document', arguments: '{"query":"gato"}' }],
          results: [{ callId: 't1', content: 'nada', isError: true }],
        },
      ],
    }
    await collect(provider.stream(prompt, new AbortController().signal))
    const body = calls[0] as { tools: unknown; messages: unknown[] }
    expect(body.tools).toEqual([
      { name: 'search_document', description: 'Busca.', input_schema: { type: 'object', properties: {} } },
    ])
    expect(body.messages.slice(1)).toEqual([
      {
        role: 'assistant',
        content: [
          { type: 'text', text: 'Vou ver.' },
          { type: 'tool_use', id: 't1', name: 'search_document', input: { query: 'gato' } },
        ],
      },
      {
        role: 'user',
        content: [{ type: 'tool_result', tool_use_id: 't1', content: 'nada', is_error: true }],
      },
    ])
  })
})
