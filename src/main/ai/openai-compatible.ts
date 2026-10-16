import type { AiModel, AiPrompt, AiProvider, AiStreamPart, AiToolExchange } from '@services/ai/provider.js'
import { OpenAiToolCalls, parseOpenAiLine } from '@services/ai/sse.js'
import { ErrorCode } from '@shared/errors.js'
import { MAX_AI_MODELS } from '@shared/limits.js'
import { aiError, errorForStatus, refusedError, toolCallsTooLargeError, unreachableError } from './errors.js'
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
      const init = { method: 'POST', headers, body, signal }
      const response = await send(fetcher, `${base}/chat/completions`, init, isFirstToolRound(prompt))
      yield* readEvents(response, signal, roundOf(prompt))
    },
  }
}

const roundOf = (prompt: AiPrompt): number => prompt.exchanges?.length ?? 0

/** Later rounds already used the tools: a 400 there is about something else. */
const isFirstToolRound = (prompt: AiPrompt): boolean =>
  (prompt.tools?.length ?? 0) > 0 && roundOf(prompt) === 0

function exchangeMessages(exchange: AiToolExchange): object[] {
  return [
    {
      role: 'assistant',
      content: exchange.text === '' ? null : exchange.text,
      tool_calls: exchange.calls.map((call) => ({
        id: call.id,
        type: 'function',
        function: { name: call.name, arguments: call.arguments },
      })),
    },
    ...exchange.results.map((result) => ({
      role: 'tool',
      tool_call_id: result.callId,
      content: result.content,
    })),
  ]
}

function toolsOf(prompt: AiPrompt): object {
  const tools = prompt.tools ?? []
  if (tools.length === 0) return {}
  return {
    tools: tools.map((tool) => ({
      type: 'function',
      function: { name: tool.name, description: tool.description, parameters: tool.parameters },
    })),
  }
}

function chatRequest(prompt: AiPrompt): object {
  return {
    model: prompt.model,
    stream: true,
    max_tokens: prompt.maxTokens,
    messages: [
      { role: 'system', content: prompt.system },
      ...prompt.messages.map((turn) => ({ role: turn.role, content: turn.text })),
      ...(prompt.exchanges ?? []).flatMap(exchangeMessages),
    ],
    ...toolsOf(prompt),
  }
}

const HTTP_BAD_REQUEST = 400
const HTTP_UNPROCESSABLE = 422
const TOOLS_MENTION = /\btools?\b/iu
/** An error body is a sentence or a small JSON; past this, the rest is not read. */
const MAX_ERROR_BODY_BYTES = 4_096

async function readErrorBody(response: Response): Promise<string> {
  if (response.body === null) return ''
  const reader = response.body.getReader()
  const decoder = new TextDecoder()
  let detail = ''
  let read = 0
  try {
    while (read < MAX_ERROR_BODY_BYTES) {
      const { done, value } = await reader.read()
      if (done) break
      const piece = value.subarray(0, MAX_ERROR_BODY_BYTES - read)
      read += piece.byteLength
      detail += decoder.decode(piece, { stream: true })
    }
  } catch {
    return detail
  } finally {
    await reader.cancel().catch(() => undefined)
  }
  return detail
}

/** Ollama answers 400 "does not support tools" for a model without them; others say much the same. */
async function refusesTools(response: Response): Promise<boolean> {
  if (response.status !== HTTP_BAD_REQUEST && response.status !== HTTP_UNPROCESSABLE) return false
  return TOOLS_MENTION.test(await readErrorBody(response))
}

async function send(
  fetcher: Fetch,
  url: string,
  init: RequestInit,
  mayRefuseTools = false,
): Promise<Response> {
  let response: Response
  try {
    response = await fetcher(url, { ...init, redirect: 'error' })
  } catch (cause) {
    if (init.signal?.aborted === true) throw cause
    throw unreachableError()
  }
  if (!response.ok) {
    if (mayRefuseTools && (await refusesTools(response)))
      throw aiError(ErrorCode.AiFailed, 'errors.ai.toolsUnsupported')
    await response.body?.cancel().catch(() => undefined)
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

async function* readEvents(
  response: Response,
  signal: AbortSignal,
  round: number,
): AsyncGenerator<AiStreamPart> {
  if (response.body === null) throw aiError(ErrorCode.AiFailed, 'errors.ai.failed')
  const reader = response.body.getReader()
  const calls = new OpenAiToolCalls(round)
  try {
    yield* readLines(reader, signal, calls)
  } finally {
    await reader.cancel().catch(() => undefined)
  }
  for (const toolCall of calls.finished()) yield { toolCall }
}

async function* readLines(
  reader: ReadableStreamDefaultReader<Uint8Array>,
  signal: AbortSignal,
  calls: OpenAiToolCalls,
): AsyncGenerator<AiStreamPart> {
  const decoder = new TextDecoder()
  let buffer = ''
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
      if (event.type === 'tools' && !calls.add(event.deltas)) throw toolCallsTooLargeError()
      if ((event.type === 'text' || event.type === 'tools') && event.text !== '') yield { text: event.text }
    }
  }
}
