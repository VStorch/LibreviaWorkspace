import { describe, expect, it } from 'vitest'
import Anthropic from '@anthropic-ai/sdk'
import { ErrorCode } from '@shared/errors.js'
import { createAnthropicProvider } from './anthropic.js'

const PROMPT = { model: 'modelo-x', system: 'sys', user: 'Olá', maxTokens: 100 }

type Client = Parameters<typeof createAnthropicProvider>[1]

function clientWith(events: unknown[], calls: unknown[] = []): Client {
  return {
    messages: {
      stream: (body: unknown, options: unknown) => {
        calls.push(body, options)
        return (async function* () {
          for (const event of events) yield event
        })()
      },
    },
    models: {
      list: () =>
        (async function* () {
          yield { id: 'modelo-a', display_name: 'Modelo A' }
          yield { id: 'modelo-b', display_name: 'Modelo B' }
        })(),
    },
  } as unknown as Client
}

async function collect(iterable: AsyncIterable<unknown>): Promise<unknown[]> {
  const parts: unknown[] = []
  for await (const part of iterable) parts.push(part)
  return parts
}

describe('provedor Anthropic', () => {
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
      messages: {
        stream: () => {
          throw new Anthropic.AuthenticationError(401, undefined, 'invalid x-api-key', new Headers())
        },
      },
    } as unknown as Client
    await expect(
      collect(createAnthropicProvider('chave', client).stream(PROMPT, new AbortController().signal)),
    ).rejects.toMatchObject({ code: ErrorCode.AiAuthFailed })
  })

  it('lista os modelos com o nome de exibição', async () => {
    const provider = createAnthropicProvider('chave', clientWith([]))
    expect(await provider.listModels(new AbortController().signal)).toEqual([
      { id: 'modelo-a', label: 'Modelo A' },
      { id: 'modelo-b', label: 'Modelo B' },
    ])
  })
})
