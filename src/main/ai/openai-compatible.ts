import type { AiModel, AiPrompt, AiProvider, AiStreamPart } from '@services/ai/provider.js'
import { parseOpenAiLine } from '@services/ai/sse.js'
import { ErrorCode } from '@shared/errors.js'
import { MAX_AI_MODELS } from '@shared/limits.js'
import { aiError, errorForStatus, refusedError, unreachableError } from './errors.js'
import { takeBoundedLines } from './stream-lines.js'

type Fetch = typeof fetch

export interface OpenAiCompatibleOptions {
  readonly baseUrl: string
  readonly apiKey?: string
  readonly fetch?: Fetch
}

interface ModelList {
  readonly data?: readonly { readonly id?: unknown }[]
}

/** Ollama, LM Studio and any server that speaks `/v1/chat/completions`. */
export function createOpenAiCompatibleProvider(options: OpenAiCompatibleOptions): AiProvider {
  const fetcher = options.fetch ?? fetch
  const base = options.baseUrl.replace(/\/+$/u, '')
  const headers: Record<string, string> = { 'content-type': 'application/json' }
  if (options.apiKey !== undefined && options.apiKey !== '')
    headers['authorization'] = `Bearer ${options.apiKey}`

  return {
    async listModels(signal): Promise<AiModel[]> {
      const response = await send(fetcher, `${base}/models`, { headers, signal })
      const body = (await response.json()) as ModelList
      return (body.data ?? [])
        .flatMap((model) =>
          typeof model.id === 'string' && model.id !== '' ? [{ id: model.id, label: model.id }] : [],
        )
        .slice(0, MAX_AI_MODELS)
    },
    async *stream(prompt, signal): AsyncGenerator<AiStreamPart> {
      const body = JSON.stringify(chatRequest(prompt))
      const response = await send(fetcher, `${base}/chat/completions`, {
        method: 'POST',
        headers,
        body,
        signal,
      })
      yield* readEvents(response, signal)
    },
  }
}

function chatRequest(prompt: AiPrompt): object {
  return {
    model: prompt.model,
    stream: true,
    max_tokens: prompt.maxTokens,
    messages: [
      { role: 'system', content: prompt.system },
      { role: 'user', content: prompt.user },
    ],
  }
}

async function send(fetcher: Fetch, url: string, init: RequestInit): Promise<Response> {
  let response: Response
  try {
    response = await fetcher(url, { ...init, redirect: 'error' })
  } catch (cause) {
    if (init.signal?.aborted === true) throw cause
    throw unreachableError()
  }
  if (!response.ok) {
    await response.body?.cancel()
    throw errorForStatus(response.status)
  }
  return response
}

async function readChunk(
  reader: ReadableStreamDefaultReader<Uint8Array>,
  signal: AbortSignal,
): Promise<ReadableStreamReadResult<Uint8Array>> {
  try {
    return await reader.read()
  } catch (cause) {
    if (signal.aborted) throw cause
    throw unreachableError()
  }
}

async function* readEvents(response: Response, signal: AbortSignal): AsyncGenerator<AiStreamPart> {
  if (response.body === null) throw aiError(ErrorCode.AiFailed, 'errors.ai.failed')
  const reader = response.body.getReader()
  const decoder = new TextDecoder()
  let buffer = ''
  try {
    for (;;) {
      const { done, value } = await readChunk(reader, signal)
      if (done) return
      const split = takeBoundedLines(buffer + decoder.decode(value, { stream: true }))
      buffer = split.rest
      for (const line of split.lines) {
        const event = parseOpenAiLine(line)
        if (event.type === 'done') return
        if (event.type === 'refused') throw refusedError()
        if (event.type === 'error') throw aiError(ErrorCode.AiFailed, 'errors.ai.failed', event.message)
        if (event.type === 'text') yield { text: event.text }
      }
    }
  } finally {
    await reader.cancel().catch(() => undefined)
  }
}
