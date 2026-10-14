import { describe, expect, it } from 'vitest'
import type { AiProvider, AiStreamPart } from '@services/ai/provider.js'
import { generate } from './assistant.js'
import { AiModelCache } from './model-cache.js'
import { AiSettings } from './settings.js'

const REQUEST = {
  requestId: '0b8f3c1e-5d2a-4c7b-9e6f-1a2b3c4d5e6f',
  providerId: 'id1',
  model: '',
  action: 'rewrite',
  text: 'Um texto.',
} as const

function settingsWith(kind: 'anthropic' | 'openai-compatible', baseUrl?: string): AiSettings {
  let data: unknown
  const settings = new AiSettings(
    { read: () => data, write: (value) => (data = value) },
    { seal: (plain) => plain, open: (sealed) => sealed },
    () => 'id1',
  )
  settings.save({
    kind,
    label: 'P',
    apiKey: 'k',
    makeDefault: false,
    revokeConsent: false,
    ...(baseUrl ? { baseUrl } : {}),
  })
  return settings
}

function providerOf(parts: AiStreamPart[], seen: { model?: string } = {}): AiProvider {
  return {
    listModels: async () => [{ id: 'primeiro', label: 'primeiro' }],
    async *stream(prompt) {
      seen.model = prompt.model
      yield* parts
    },
  }
}

describe('generate', () => {
  it('pergunta antes de mandar para fora e não manda nada se a pessoa recusar', async () => {
    let streamed = false
    const result = await generate(REQUEST, {
      settings: settingsWith('anthropic'),
      signal: new AbortController().signal,
      push: () => undefined,
      confirmConsent: async () => false,
      providerFor: async () => {
        streamed = true
        return providerOf([])
      },
    })
    expect(result).toEqual({ status: 'declined' })
    expect(streamed).toBe(false)
  })

  it('lembra o consentimento e usa o primeiro modelo quando nenhum foi escolhido', async () => {
    const settings = settingsWith('anthropic')
    const pushed: string[] = []
    const seen: { model?: string } = {}
    const result = await generate(REQUEST, {
      settings,
      signal: new AbortController().signal,
      push: (delta) => pushed.push(delta),
      confirmConsent: async () => true,
      providerFor: async () => providerOf([{ text: 'Novo ' }, { text: 'texto.' }, { costUsd: 0.5 }], seen),
    })
    expect(result).toEqual({ status: 'done', model: 'primeiro', costUsd: 0.5 })
    expect(pushed.join('')).toBe('Novo texto.')
    expect(seen.model).toBe('primeiro')
    expect(settings.list().providers[0]?.consented).toBe(true)
  })

  it('servidor local não pede consentimento', async () => {
    let asked = false
    const result = await generate(
      { ...REQUEST, model: 'llama3' },
      {
        settings: settingsWith('openai-compatible', 'http://127.0.0.1:11434/v1'),
        signal: new AbortController().signal,
        push: () => undefined,
        confirmConsent: async () => (asked = true),
        providerFor: async () => providerOf([{ text: 'ok' }]),
      },
    )
    expect(result).toEqual({ status: 'done', model: 'llama3' })
    expect(asked).toBe(false)
  })

  it('cancelado devolve o estado, não um erro', async () => {
    const controller = new AbortController()
    const provider: AiProvider = {
      listModels: async () => [],
      async *stream(_prompt, signal) {
        yield { text: 'um' }
        controller.abort()
        signal.throwIfAborted()
      },
    }
    const result = await generate(
      { ...REQUEST, model: 'm' },
      {
        settings: settingsWith('openai-compatible', 'http://localhost:1/v1'),
        signal: controller.signal,
        push: () => undefined,
        confirmConsent: async () => true,
        providerFor: async () => provider,
      },
    )
    expect(result).toEqual({ status: 'canceled' })
  })

  it('o modelo informado pelo fluxo vira o autor, mesmo no automático', async () => {
    const result = await generate(REQUEST, {
      settings: settingsWith('openai-compatible', 'http://localhost:1/v1'),
      signal: new AbortController().signal,
      push: () => undefined,
      confirmConsent: async () => true,
      providerFor: async () => providerOf([{ model: 'claude-opus-5-5' }, { text: 'ok' }]),
    })
    expect(result).toEqual({ status: 'done', model: 'claude-opus-5-5' })
  })

  it('no automático, a lista de modelos é buscada uma vez só', async () => {
    let listed = 0
    const provider: AiProvider = {
      listModels: async () => {
        listed += 1
        return [{ id: 'primeiro', label: 'primeiro' }]
      },
      async *stream() {
        yield { text: 'ok' }
      },
    }
    const context = {
      settings: settingsWith('openai-compatible', 'http://localhost:1/v1'),
      signal: new AbortController().signal,
      push: () => undefined,
      confirmConsent: async () => true,
      providerFor: async () => provider,
      models: new AiModelCache(),
    }
    await generate(REQUEST, context)
    await generate({ ...REQUEST, requestId: '1b8f3c1e-5d2a-4c7b-9e6f-1a2b3c4d5e6f' }, context)
    expect(listed).toBe(1)
    context.models.forget('id1')
    await generate(REQUEST, context)
    expect(listed).toBe(2)
  })
})
