import { describe, expect, it } from 'vitest'
import type { AiProvider, AiStreamPart } from '@services/ai/provider.js'
import { DEFAULT_LOCAL_AI_CONTEXT_LENGTH } from '@shared/limits.js'
import { consentDestination, generate } from './assistant.js'
import { AiModelCache } from './model-cache.js'
import { AiSettings, type StoredAiProvider } from './settings.js'

const REQUEST = {
  requestId: '0b8f3c1e-5d2a-4c7b-9e6f-1a2b3c4d5e6f',
  providerId: 'id1',
  model: '',
  action: 'rewrite',
  text: 'Um texto.',
  before: '',
  after: '',
  format: 'plain',
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

  it('servidor local recebe só o documento em volta da seleção que cabe no limite dele', async () => {
    let user = ''
    const provider: AiProvider = {
      listModels: async () => [],
      async *stream(prompt) {
        user = prompt.user
        yield { text: 'ok' }
      },
    }
    await generate(
      { ...REQUEST, model: 'llama3', before: 'a '.repeat(5_000), after: ' depois' },
      {
        settings: settingsWith('openai-compatible', 'http://127.0.0.1:11434/v1'),
        signal: new AbortController().signal,
        push: () => undefined,
        confirmConsent: async () => true,
        providerFor: async () => provider,
      },
    )
    expect(user).toContain('<selection>Um texto.</selection> depois')
    expect(user.length).toBeLessThanOrEqual(DEFAULT_LOCAL_AI_CONTEXT_LENGTH + 100)
  })

  it('agente ACP pede consentimento e, no automático, usa o modelo do próprio agente sem abrir outro processo para listar', async () => {
    let data: unknown
    const settings = new AiSettings(
      { read: () => data, write: (value) => (data = value) },
      { seal: (plain) => plain, open: (sealed) => sealed },
      () => 'id1',
    )
    settings.save({
      kind: 'acp',
      label: 'Gemini',
      agentId: 'gemini',
      makeDefault: false,
      revokeConsent: false,
    })
    let asked = false
    let listed = false
    const seen: { model?: string } = {}
    const result = await generate(REQUEST, {
      settings,
      signal: new AbortController().signal,
      push: () => undefined,
      confirmConsent: async () => (asked = true),
      providerFor: async () => ({
        ...providerOf([{ model: 'gemini-pro' }, { text: 'ok' }], seen),
        listModels: async () => {
          listed = true
          return []
        },
      }),
    })
    expect(asked).toBe(true)
    expect(listed).toBe(false)
    expect(seen.model).toBe('')
    expect(result).toEqual({ status: 'done', model: 'gemini-pro' })
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

  describe('troca de modelo quando a Anthropic recusa', () => {
    const fallbackProvider = (seen: { fallback?: boolean }): AiProvider => ({
      listModels: async () => [
        { id: 'aceita', label: 'Aceita', fallback: true },
        { id: 'nao-aceita', label: 'Não aceita', fallback: false },
      ],
      async *stream(prompt) {
        seen.fallback = prompt.fallback === true
        yield { text: 'Parte. ' }
        yield { model: 'substituto' }
        yield { text: 'Resto.' }
      },
    })
    const run = async (
      model: string,
      preference: boolean,
    ): Promise<{ fallback?: boolean; author?: string | undefined }> => {
      const settings = settingsWith('anthropic')
      settings.markConsented('id1')
      settings.savePreferences({ keepFormatting: true, anthropicFallback: preference })
      const seen: { fallback?: boolean } = {}
      const result = await generate(
        { ...REQUEST, model },
        {
          settings,
          signal: new AbortController().signal,
          push: () => undefined,
          confirmConsent: async () => true,
          providerFor: async () => fallbackProvider(seen),
        },
      )
      return { ...seen, ...(result.status === 'done' ? { author: result.model } : {}) }
    }

    it('com a preferência ligada e um modelo que aceita, pede a troca e o autor é o modelo final', async () => {
      expect(await run('aceita', true)).toEqual({ fallback: true, author: 'substituto' })
    })

    it('não pede a troca quando o modelo não aceita ou a preferência está desligada', async () => {
      expect((await run('nao-aceita', true)).fallback).toBe(false)
      expect((await run('aceita', false)).fallback).toBe(false)
    })
  })
})

describe('consentDestination', () => {
  it('para um agente ACP mostra o programa que vai rodar e os argumentos, não o nome dado', () => {
    const chosen = { id: 'a', kind: 'acp', label: 'Seguro', command: '/opt/x/agente', args: ['--acp', 'a b'] }
    expect(consentDestination(chosen as StoredAiProvider)).toBe("/opt/x/agente --acp 'a b'")
    const known = { id: 'b', kind: 'acp', label: 'Gemini', agentId: 'gemini', args: ['--acp'] }
    expect(consentDestination(known as StoredAiProvider)).toBe('gemini --acp')
  })

  it('para um servidor mostra o endereço', () => {
    const server = { id: 'c', kind: 'openai-compatible', label: 'S', baseUrl: 'https://a.example/v1' }
    expect(consentDestination(server as StoredAiProvider)).toBe('https://a.example/v1')
  })
})
