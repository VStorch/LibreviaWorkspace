import { describe, expect, it } from 'vitest'
import type { AiPrompt, AiProvider, AiStreamPart } from '@services/ai/provider.js'
import { DEFAULT_LOCAL_AI_CONTEXT_LENGTH } from '@shared/limits.js'
import { chat, consentDestination, generate, sheet, type AiConversationRequest } from './assistant.js'
import { AiModelCache } from './model-cache.js'
import { AiSettings, type StoredAiProvider } from './settings.js'
import type { ToolSessions } from './tool-server.js'

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

function settingsWith(kind: 'anthropic' | 'openai-compatible' | 'claude-cli', baseUrl?: string): AiSettings {
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
    listModels: async () => [{ id: 'first', label: 'first' }],
    async *stream(prompt) {
      seen.model = prompt.model
      yield* parts
    },
  }
}

describe('generate', () => {
  it('asks before sending out and sends nothing if the person declines', async () => {
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

  it('remembers the consent and uses the first model when none was chosen', async () => {
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
    expect(result).toEqual({ status: 'done', model: 'first', costUsd: 0.5 })
    expect(pushed.join('')).toBe('Novo texto.')
    expect(seen.model).toBe('first')
    expect(settings.list().providers[0]?.consented).toBe(true)
  })

  it('a local server does not ask for consent', async () => {
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

  it('a local server gets only the document around the selection that fits its limit', async () => {
    let user = ''
    const provider: AiProvider = {
      listModels: async () => [],
      async *stream(prompt) {
        user = prompt.messages[0]?.text ?? ''
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

  it("an ACP agent asks for consent and, in automatic mode, uses the agent's own model without opening another process to list", async () => {
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

  it('cancelled returns the state, not an error', async () => {
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

  it('the model reported by the stream becomes the author, even in automatic mode', async () => {
    const result = await generate(REQUEST, {
      settings: settingsWith('openai-compatible', 'http://localhost:1/v1'),
      signal: new AbortController().signal,
      push: () => undefined,
      confirmConsent: async () => true,
      providerFor: async () => providerOf([{ model: 'claude-opus-5-5' }, { text: 'ok' }]),
    })
    expect(result).toEqual({ status: 'done', model: 'claude-opus-5-5' })
  })

  it('in automatic mode, the model list is fetched only once', async () => {
    let listed = 0
    const provider: AiProvider = {
      listModels: async () => {
        listed += 1
        return [{ id: 'first', label: 'first' }]
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

  describe('model switch when Anthropic refuses', () => {
    const fallbackProvider = (seen: { fallback?: boolean }): AiProvider => ({
      listModels: async () => [
        { id: 'aceita', label: 'Aceita', fallback: true },
        { id: 'not-accepted', label: 'Not accepted', fallback: false },
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
      settings.savePreferences({
        keepFormatting: true,
        anthropicFallback: preference,
        chatProposals: false,
        mcpServer: false,
      })
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

    it('with the preference on and a model that accepts, asks for the switch and the author is the final model', async () => {
      expect(await run('aceita', true)).toEqual({ fallback: true, author: 'substituto' })
    })

    it('does not ask for the switch when the model does not accept it or the preference is off', async () => {
      expect((await run('not-accepted', true)).fallback).toBe(false)
      expect((await run('aceita', false)).fallback).toBe(false)
    })
  })
})

describe('consentDestination', () => {
  it('for an ACP agent shows the program that will run and its arguments, not the given name', () => {
    const chosen = { id: 'a', kind: 'acp', label: 'Seguro', command: '/opt/x/agente', args: ['--acp', 'a b'] }
    expect(consentDestination(chosen as StoredAiProvider)).toBe("/opt/x/agente --acp 'a b'")
    const known = { id: 'b', kind: 'acp', label: 'Gemini', agentId: 'gemini', args: ['--acp'] }
    expect(consentDestination(known as StoredAiProvider)).toBe('gemini --acp')
  })

  it('for a server shows the address', () => {
    const server = { id: 'c', kind: 'openai-compatible', label: 'S', baseUrl: 'https://a.example/v1' }
    expect(consentDestination(server as StoredAiProvider)).toBe('https://a.example/v1')
  })
})

const CHAT: AiConversationRequest = {
  requestId: '1c9a4d2f-6e3b-4d8c-af70-2b3c4d5e6f70',
  providerId: 'id1',
  model: 'm',
  messages: [
    { role: 'user', text: 'Qual é o tema?' },
    { role: 'assistant', text: 'É sobre chuva.' },
    { role: 'user', text: 'E o tom?' },
  ],
  before: 'Choveu o dia todo.',
  after: '',
  selection: '',
  proposals: false,
  paragraphs: [],
}

function recordingProvider(parts: AiStreamPart[], seen: { prompt?: AiPrompt }): AiProvider {
  return {
    listModels: async () => [],
    async *stream(prompt) {
      seen.prompt = prompt
      yield* parts
    },
  }
}

describe('chat', () => {
  it('sends the history as messages and the current document in the system prompt, and returns the text as a stream', async () => {
    const seen: { prompt?: AiPrompt } = {}
    const pushed: string[] = []
    const result = await chat(CHAT, {
      settings: settingsWith('openai-compatible', 'http://127.0.0.1:11434/v1'),
      signal: new AbortController().signal,
      push: (delta) => pushed.push(delta),
      confirmConsent: async () => true,
      providerFor: async () => recordingProvider([{ text: 'Sereno' }, { text: '.' }], seen),
    })
    expect(result).toEqual({ status: 'done', model: 'm' })
    expect(pushed.join('')).toBe('Sereno.')
    expect(seen.prompt?.messages).toEqual(CHAT.messages)
    expect(seen.prompt?.system).toContain('Choveu o dia todo.')
  })

  it('sends nothing if the person declines consent', async () => {
    const seen: { prompt?: AiPrompt } = {}
    const result = await chat(CHAT, {
      settings: settingsWith('anthropic'),
      signal: new AbortController().signal,
      push: () => undefined,
      confirmConsent: async () => false,
      providerFor: async () => recordingProvider([{ text: 'x' }], seen),
    })
    expect(result).toEqual({ status: 'declined' })
    expect(seen.prompt).toBeUndefined()
  })

  it('cancelled midway returns the state, not an error', async () => {
    const controller = new AbortController()
    const provider: AiProvider = {
      listModels: async () => [],
      async *stream(_prompt, signal) {
        yield { text: 'meio' }
        controller.abort()
        signal.throwIfAborted()
      },
    }
    const pushed: string[] = []
    const result = await chat(CHAT, {
      settings: settingsWith('openai-compatible', 'http://127.0.0.1:11434/v1'),
      signal: controller.signal,
      push: (delta) => pushed.push(delta),
      confirmConsent: async () => true,
      providerFor: async () => provider,
    })
    expect(result).toEqual({ status: 'canceled' })
    expect(pushed.join('')).toBe('meio')
  })

  it('with proposals on and a compatible server, sends the tools and returns the proposals', async () => {
    const settings = settingsWith('openai-compatible', 'http://127.0.0.1:11434/v1')
    settings.savePreferences({
      keepFormatting: true,
      anthropicFallback: false,
      chatProposals: true,
      mcpServer: false,
    })
    const prompts: AiPrompt[] = []
    const provider: AiProvider = {
      listModels: async () => [],
      async *stream(prompt) {
        prompts.push(prompt)
        if (prompts.length === 1) {
          yield {
            toolCall: {
              id: 'a',
              name: 'propose_edit',
              arguments: '{"paragraph":1,"find":"dia","replace":"mês"}',
            },
          }
        } else yield { text: 'Proposto.' }
      },
    }
    const result = await chat(
      { ...CHAT, proposals: true, paragraphs: ['Choveu o dia todo.'] },
      {
        settings,
        signal: new AbortController().signal,
        push: () => undefined,
        confirmConsent: async () => true,
        providerFor: async () => provider,
      },
    )
    expect(result).toEqual({
      status: 'done',
      model: 'm',
      proposals: [
        {
          id: '1',
          kind: 'edit',
          paragraph: 1,
          paragraphText: 'Choveu o dia todo.',
          find: 'dia',
          replace: 'mês',
        },
      ],
    })
    expect(prompts[0]?.system).toContain('[1] Choveu o dia todo.')
    expect(prompts[0]?.tools).toHaveLength(5)
  })

  it('with Claude Code, the tools go through the local MCP server, in a session that closes with the response', async () => {
    const settings = settingsWith('claude-cli')
    settings.savePreferences({
      keepFormatting: true,
      anthropicFallback: false,
      chatProposals: true,
      mcpServer: false,
    })
    const proposal = {
      id: '1',
      kind: 'comment' as const,
      paragraph: 1,
      paragraphText: 'Choveu.',
      find: 'Choveu',
      comment: 'Bom.',
    }
    const events: string[] = []
    const opened: unknown[] = []
    const toolSessions: ToolSessions = {
      open: async (document) => {
        opened.push(document)
        return {
          server: { name: 'librevia', command: 'x', args: [], env: {}, tools: ['add_comment'] },
          proposals: () => [proposal],
          close: () => events.push('fechou'),
        }
      },
    }
    const seen: { prompt?: AiPrompt } = {}
    const result = await chat(
      { ...CHAT, proposals: true, paragraphs: ['Choveu.'] },
      {
        settings,
        signal: new AbortController().signal,
        push: () => undefined,
        confirmConsent: async () => true,
        providerFor: async () => recordingProvider([{ text: 'Comentei.' }], seen),
        toolSessions,
      },
    )
    expect(result).toMatchObject({ status: 'done', proposals: [proposal] })
    expect(seen.prompt?.toolServer?.tools).toEqual(['add_comment'])
    expect(seen.prompt?.tools).toBeUndefined()
    expect(opened).toEqual([{ paragraphs: ['Choveu.'], readBudget: expect.any(Number) }])
    expect(events).toEqual(['fechou'])
  })

  it('with proposals off, the request carries no tools', async () => {
    const seen: { prompt?: AiPrompt } = {}
    const result = await chat(
      { ...CHAT, paragraphs: ['Choveu o dia todo.'] },
      {
        settings: settingsWith('openai-compatible', 'http://127.0.0.1:11434/v1'),
        signal: new AbortController().signal,
        push: () => undefined,
        confirmConsent: async () => true,
        providerFor: async () => recordingProvider([{ text: 'ok' }], seen),
      },
    )
    expect(result).toEqual({ status: 'done', model: 'm' })
    expect(seen.prompt?.tools).toBeUndefined()
    expect(seen.prompt?.system).not.toContain('[1]')
  })

  it('with the preference on but the request without the mark, sends no tools', async () => {
    const settings = settingsWith('openai-compatible', 'http://127.0.0.1:11434/v1')
    settings.savePreferences({
      keepFormatting: true,
      anthropicFallback: false,
      chatProposals: true,
      mcpServer: false,
    })
    const seen: { prompt?: AiPrompt } = {}
    await chat(
      { ...CHAT, paragraphs: ['Choveu o dia todo.'] },
      {
        settings,
        signal: new AbortController().signal,
        push: () => undefined,
        confirmConsent: async () => true,
        providerFor: async () => recordingProvider([{ text: 'ok' }], seen),
      },
    )
    expect(seen.prompt?.tools).toBeUndefined()
  })
})

describe('sheet', () => {
  const SHEET = {
    requestId: REQUEST.requestId,
    providerId: 'id1',
    model: 'm',
    task: 'formula',
    language: 'pt',
    instruction: 'some os preços',
    context: 'Sheet: Vendas\nActive cell: B4\nCells:\nB2: 3\nB3: 5',
  } as const

  function contextOf(stream: AiStreamPart[], seen: { prompt?: AiPrompt }, pushed: string[] = []) {
    return {
      settings: settingsWith('openai-compatible', 'http://127.0.0.1:11434/v1'),
      signal: new AbortController().signal,
      push: (delta: string) => pushed.push(delta),
      confirmConsent: async () => true,
      providerFor: async () => recordingProvider(stream, seen),
    }
  }

  it('sends the sheet as context, without tools, and returns the response as a stream', async () => {
    const seen: { prompt?: AiPrompt } = {}
    const pushed: string[] = []
    const result = await sheet(SHEET, contextOf([{ text: '=SOMA(' }, { text: 'B2:B3)' }], seen, pushed))
    expect(result).toEqual({ status: 'done', model: 'm' })
    expect(pushed.join('')).toBe('=SOMA(B2:B3)')
    expect(seen.prompt?.tools).toBeUndefined()
    expect(seen.prompt?.system).toContain('SOMA')
    expect(seen.prompt?.messages[0]?.text).toContain('B2: 3')
  })

  it('rows to fill that do not fit the local provider context give an error without sending', async () => {
    const seen: { prompt?: AiPrompt } = {}
    const fill = {
      ...SHEET,
      task: 'fill',
      rowCount: 1,
      context: 'x'.repeat(DEFAULT_LOCAL_AI_CONTEXT_LENGTH),
    } as const
    await expect(sheet(fill, contextOf([{ text: '[1]' }], seen))).rejects.toThrow()
    expect(seen.prompt).toBeUndefined()
  })
})
