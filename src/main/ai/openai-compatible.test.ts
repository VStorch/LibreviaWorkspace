import { describe, expect, it, vi } from 'vitest'
import { ErrorCode } from '@shared/errors.js'
import { MAX_AI_MODELS, MAX_AI_STREAM_LINE_LENGTH, MAX_AI_TOOL_CALLS } from '@shared/limits.js'
import { t } from '../i18n.js'
import { createOpenAiCompatibleProvider } from './openai-compatible.js'

const PROMPT = {
  model: 'llama3',
  system: 'sys',
  messages: [{ role: 'user' as const, text: 'Olá' }],
  maxTokens: 100,
}

function sseResponse(lines: string[]): Response {
  const encoder = new TextEncoder()
  const body = new ReadableStream<Uint8Array>({
    start(controller) {
      for (const line of lines) controller.enqueue(encoder.encode(line))
      controller.close()
    },
  })
  return new Response(body, { status: 200, headers: { 'content-type': 'text/event-stream' } })
}

const delta = (content: string): string =>
  `data: ${JSON.stringify({ choices: [{ delta: { content } }] })}\n\n`

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
describe('OpenAI-compatible provider', () => {
  it('sends the whole conversation, after the system prompt, with the role of each turn', async () => {
    const fetchStub = vi.fn(async () => sseResponse(['data: [DONE]\n\n']))
    const provider = createOpenAiCompatibleProvider({
      baseUrl: 'http://localhost:11434/v1',
      fetch: fetchStub as unknown as typeof fetch,
    })
    await collect(provider.stream(CONVERSATION, new AbortController().signal))
    const [, init] = fetchStub.mock.calls[0] as unknown as [string, RequestInit]
    expect(JSON.parse(init.body as string).messages).toEqual([
      { role: 'system', content: 'sys' },
      { role: 'user', content: 'Qual o tema?' },
      { role: 'assistant', content: 'Chuva.' },
      { role: 'user', content: 'E o tom?' },
    ])
  })

  it('reads the text in chunks, even with a line split across two packets', async () => {
    const half = delta(' mundo')
    const fetchStub = vi.fn(async () =>
      sseResponse([delta('Olá'), half.slice(0, 20), half.slice(20), 'data: [DONE]\n\n', delta('ignorado')]),
    )
    const provider = createOpenAiCompatibleProvider({
      baseUrl: 'http://localhost:11434/v1/',
      apiKey: 'segredo',
      fetch: fetchStub as unknown as typeof fetch,
    })

    expect(await collect(provider.stream(PROMPT, new AbortController().signal))).toEqual([
      { text: 'Olá' },
      { text: ' mundo' },
    ])
    const [url, init] = fetchStub.mock.calls[0] as unknown as [string, RequestInit]
    expect(url).toBe('http://localhost:11434/v1/chat/completions')
    expect(init.redirect).toBe('error')
    expect((init.headers as Record<string, string>)['authorization']).toBe('Bearer segredo')
    expect(JSON.parse(init.body as string)).toMatchObject({ model: 'llama3', stream: true, max_tokens: 100 })
  })

  it('translates 401 into a key error', async () => {
    const provider = createOpenAiCompatibleProvider({
      baseUrl: 'http://localhost:1/v1',
      fetch: (async () => new Response('no', { status: 401 })) as unknown as typeof fetch,
    })
    await expect(collect(provider.stream(PROMPT, new AbortController().signal))).rejects.toMatchObject({
      code: ErrorCode.AiAuthFailed,
    })
  })

  it('a server that is down becomes a readable error', async () => {
    const provider = createOpenAiCompatibleProvider({
      baseUrl: 'http://localhost:1/v1',
      fetch: (async () => {
        throw new TypeError('fetch failed')
      }) as unknown as typeof fetch,
    })
    await expect(provider.listModels(new AbortController().signal)).rejects.toMatchObject({
      code: ErrorCode.AiUnreachable,
    })
  })

  it('cancelling closes the stream', async () => {
    const controller = new AbortController()
    let canceled = false
    const body = new ReadableStream<Uint8Array>({
      start(stream) {
        stream.enqueue(new TextEncoder().encode(delta('um')))
      },
      cancel() {
        canceled = true
      },
    })
    const provider = createOpenAiCompatibleProvider({
      baseUrl: 'http://localhost:1/v1',
      fetch: (async () => new Response(body)) as unknown as typeof fetch,
    })
    for await (const part of provider.stream(PROMPT, controller.signal)) {
      expect(part).toEqual({ text: 'um' })
      controller.abort()
      break
    }
    expect(canceled).toBe(true)
  })

  it('lists the models from /models', async () => {
    const provider = createOpenAiCompatibleProvider({
      baseUrl: 'http://localhost:1/v1',
      fetch: (async () =>
        Response.json({ data: [{ id: 'llama3' }, { id: '' }, { nome: 'x' }] })) as unknown as typeof fetch,
    })
    expect(await provider.listModels(new AbortController().signal)).toEqual([
      { id: 'llama3', label: 'llama3' },
    ])
  })

  it('a connection that drops mid-response becomes a network error', async () => {
    const body = new ReadableStream<Uint8Array>({
      start(stream) {
        stream.enqueue(new TextEncoder().encode(delta('um')))
      },
      pull(stream) {
        stream.error(new TypeError('terminated'))
      },
    })
    const provider = createOpenAiCompatibleProvider({
      baseUrl: 'http://localhost:1/v1',
      fetch: (async () => new Response(body)) as unknown as typeof fetch,
    })
    await expect(collect(provider.stream(PROMPT, new AbortController().signal))).rejects.toMatchObject({
      code: ErrorCode.AiUnreachable,
    })
  })

  it('an endless line over the limit becomes a readable error', async () => {
    const provider = createOpenAiCompatibleProvider({
      baseUrl: 'http://localhost:1/v1',
      fetch: (async () =>
        sseResponse([`data: ${'x'.repeat(MAX_AI_STREAM_LINE_LENGTH + 1)}`])) as unknown as typeof fetch,
    })
    await expect(collect(provider.stream(PROMPT, new AbortController().signal))).rejects.toMatchObject({
      code: ErrorCode.AiFailed,
    })
  })

  it('content filter becomes a refusal', async () => {
    const provider = createOpenAiCompatibleProvider({
      baseUrl: 'http://localhost:1/v1',
      fetch: (async () =>
        sseResponse([
          'data: {"choices":[{"delta":{},"finish_reason":"content_filter"}]}\n\n',
        ])) as unknown as typeof fetch,
    })
    await expect(collect(provider.stream(PROMPT, new AbortController().signal))).rejects.toMatchObject({
      code: ErrorCode.AiRefused,
    })
  })

  it('lists at most the model limit', async () => {
    const data = Array.from({ length: MAX_AI_MODELS + 5 }, (_, index) => ({ id: `m${index}` }))
    const provider = createOpenAiCompatibleProvider({
      baseUrl: 'http://localhost:1/v1',
      fetch: (async () => Response.json({ data })) as unknown as typeof fetch,
    })
    expect(await provider.listModels(new AbortController().signal)).toHaveLength(MAX_AI_MODELS)
  })

  const TOOLS = [{ name: 'search_document', description: 'Busca.', parameters: { type: 'object' } }]

  it('with tools, sends tools and the history of calls and results in the OpenAI format', async () => {
    const fetchStub = vi.fn(async () => sseResponse(['data: [DONE]\n\n']))
    const provider = createOpenAiCompatibleProvider({
      baseUrl: 'http://x/v1',
      fetch: fetchStub as unknown as typeof fetch,
    })
    const exchanges = [
      {
        text: '',
        calls: [{ id: 'a', name: 'search_document', arguments: '{"query":"gato"}' }],
        results: [{ callId: 'a', content: '[1] O gato.', isError: false }],
      },
    ]
    await collect(provider.stream({ ...PROMPT, tools: TOOLS, exchanges }, new AbortController().signal))
    const [, init] = fetchStub.mock.calls[0] as unknown as [string, RequestInit]
    const body = JSON.parse(init.body as string)
    expect(body.tools).toEqual([
      {
        type: 'function',
        function: { name: 'search_document', description: 'Busca.', parameters: { type: 'object' } },
      },
    ])
    expect(body.messages.slice(2)).toEqual([
      {
        role: 'assistant',
        content: null,
        tool_calls: [
          { id: 'a', type: 'function', function: { name: 'search_document', arguments: '{"query":"gato"}' } },
        ],
      },
      { role: 'tool', tool_call_id: 'a', content: '[1] O gato.' },
    ])
  })

  it('without tools, the request carries no tools', async () => {
    const fetchStub = vi.fn(async () => sseResponse(['data: [DONE]\n\n']))
    const provider = createOpenAiCompatibleProvider({
      baseUrl: 'http://x/v1',
      fetch: fetchStub as unknown as typeof fetch,
    })
    await collect(provider.stream(PROMPT, new AbortController().signal))
    const [, init] = fetchStub.mock.calls[0] as unknown as [string, RequestInit]
    expect(JSON.parse(init.body as string)).not.toHaveProperty('tools')
  })

  it('joins the tool_calls deltas and delivers the calls at the end of the stream', async () => {
    const piece = (calls: unknown[]): string =>
      `data: ${JSON.stringify({ choices: [{ delta: { tool_calls: calls } }] })}\n\n`
    const provider = createOpenAiCompatibleProvider({
      baseUrl: 'http://x/v1',
      fetch: (async () =>
        sseResponse([
          delta('Vou ver. '),
          piece([{ index: 0, id: 'a', function: { name: 'search_document', arguments: '{"que' } }]),
          piece([{ index: 0, function: { arguments: 'ry":"gato"}' } }]),
          'data: [DONE]\n\n',
        ])) as unknown as typeof fetch,
    })
    expect(await collect(provider.stream({ ...PROMPT, tools: TOOLS }, new AbortController().signal))).toEqual(
      [
        { text: 'Vou ver. ' },
        { toolCall: { id: 'a', name: 'search_document', arguments: '{"query":"gato"}' } },
      ],
    )
  })

  it('a model without tools rejected by the server becomes a notice to turn off the checkbox', async () => {
    const refusal = JSON.stringify({
      error: { message: 'registry.ollama.ai/library/gemma does not support tools' },
    })
    const provider = createOpenAiCompatibleProvider({
      baseUrl: 'http://x/v1',
      fetch: (async () => new Response(refusal, { status: 400 })) as unknown as typeof fetch,
    })
    await expect(
      collect(provider.stream({ ...PROMPT, tools: TOOLS }, new AbortController().signal)),
    ).rejects.toMatchObject({ code: ErrorCode.AiFailed, message: t('errors.ai.toolsUnsupported') })
  })

  it('a 400 that mentions tools after the first round stays an ordinary error', async () => {
    const provider = createOpenAiCompatibleProvider({
      baseUrl: 'http://x/v1',
      fetch: (async () => new Response('bad tool_call_id', { status: 400 })) as unknown as typeof fetch,
    })
    const exchanges = [{ text: '', calls: [], results: [] }]
    await expect(
      collect(provider.stream({ ...PROMPT, tools: TOOLS, exchanges }, new AbortController().signal)),
    ).rejects.toMatchObject({ message: t('errors.ai.failed') })
  })

  it('reads only the start of a huge error body', async () => {
    const huge = 'x'.repeat(1_000_000) + ' tools'
    const provider = createOpenAiCompatibleProvider({
      baseUrl: 'http://x/v1',
      fetch: (async () => new Response(huge, { status: 400 })) as unknown as typeof fetch,
    })
    await expect(
      collect(provider.stream({ ...PROMPT, tools: TOOLS }, new AbortController().signal)),
    ).rejects.toMatchObject({ message: t('errors.ai.failed') })
  })

  it('aborts a response that asks for too many tools', async () => {
    const calls = Array.from({ length: MAX_AI_TOOL_CALLS + 1 }, (_, index) => ({
      index,
      id: `c${index}`,
      function: { name: 'search_document', arguments: '{}' },
    }))
    const provider = createOpenAiCompatibleProvider({
      baseUrl: 'http://x/v1',
      fetch: (async () =>
        sseResponse([
          `data: ${JSON.stringify({ choices: [{ delta: { tool_calls: calls } }] })}\n\n`,
        ])) as unknown as typeof fetch,
    })
    await expect(
      collect(provider.stream({ ...PROMPT, tools: TOOLS }, new AbortController().signal)),
    ).rejects.toMatchObject({ message: t('errors.ai.toolCallsTooLarge') })
  })
})
