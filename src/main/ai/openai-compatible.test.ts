import { describe, expect, it, vi } from 'vitest'
import { ErrorCode } from '@shared/errors.js'
import { MAX_AI_MODELS, MAX_AI_STREAM_LINE_LENGTH } from '@shared/limits.js'
import { createOpenAiCompatibleProvider } from './openai-compatible.js'

const PROMPT = { model: 'llama3', system: 'sys', user: 'Olá', maxTokens: 100 }

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

describe('provedor compatível com OpenAI', () => {
  it('lê o texto em pedaços, mesmo com a linha partida entre dois pacotes', async () => {
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

  it('traduz 401 em erro de chave', async () => {
    const provider = createOpenAiCompatibleProvider({
      baseUrl: 'http://localhost:1/v1',
      fetch: (async () => new Response('no', { status: 401 })) as unknown as typeof fetch,
    })
    await expect(collect(provider.stream(PROMPT, new AbortController().signal))).rejects.toMatchObject({
      code: ErrorCode.AiAuthFailed,
    })
  })

  it('servidor fora do ar vira erro legível', async () => {
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

  it('cancelar fecha o fluxo', async () => {
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

  it('lista os modelos de /models', async () => {
    const provider = createOpenAiCompatibleProvider({
      baseUrl: 'http://localhost:1/v1',
      fetch: (async () =>
        Response.json({ data: [{ id: 'llama3' }, { id: '' }, { nome: 'x' }] })) as unknown as typeof fetch,
    })
    expect(await provider.listModels(new AbortController().signal)).toEqual([
      { id: 'llama3', label: 'llama3' },
    ])
  })

  it('conexão que cai no meio da resposta vira erro de rede', async () => {
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

  it('linha sem fim maior que o limite vira erro legível', async () => {
    const provider = createOpenAiCompatibleProvider({
      baseUrl: 'http://localhost:1/v1',
      fetch: (async () =>
        sseResponse([`data: ${'x'.repeat(MAX_AI_STREAM_LINE_LENGTH + 1)}`])) as unknown as typeof fetch,
    })
    await expect(collect(provider.stream(PROMPT, new AbortController().signal))).rejects.toMatchObject({
      code: ErrorCode.AiFailed,
    })
  })

  it('filtro de conteúdo vira recusa', async () => {
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

  it('lista no máximo o limite de modelos', async () => {
    const data = Array.from({ length: MAX_AI_MODELS + 5 }, (_, index) => ({ id: `m${index}` }))
    const provider = createOpenAiCompatibleProvider({
      baseUrl: 'http://localhost:1/v1',
      fetch: (async () => Response.json({ data })) as unknown as typeof fetch,
    })
    expect(await provider.listModels(new AbortController().signal)).toHaveLength(MAX_AI_MODELS)
  })
})
