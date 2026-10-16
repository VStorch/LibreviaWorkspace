import Anthropic from '@anthropic-ai/sdk'
import type {
  BetaMessageParam,
  BetaRawMessageStreamEvent,
  BetaTool,
} from '@anthropic-ai/sdk/resources/beta/messages/messages.js'
import type {
  AiModel,
  AiPrompt,
  AiProvider,
  AiStreamPart,
  AiToolCall,
  AiToolExchange,
} from '@services/ai/provider.js'
import { ErrorCode } from '@shared/errors.js'
import { MAX_AI_MODELS, MAX_AI_TOOL_ARGUMENTS_LENGTH, MAX_AI_TOOL_CALLS } from '@shared/limits.js'
import { aiError, refusedError, toolCallsTooLargeError, unreachableError } from './errors.js'

type AnthropicClient = { readonly beta: Pick<Anthropic['beta'], 'models' | 'messages'> }

/** The listing beta reports which models accept `fallbacks`; the messages beta accepts it. */
const FALLBACK_LISTING_BETA = 'server-side-fallback-2026-06-01'
const FALLBACK_MESSAGES_BETA = 'server-side-fallback-2026-07-01'

export function createAnthropicProvider(
  apiKey: string,
  client: AnthropicClient = new Anthropic({ apiKey }),
): AiProvider {
  return {
    async listModels(signal): Promise<AiModel[]> {
      try {
        const models: AiModel[] = []
        const listing = client.beta.models.list({ betas: [FALLBACK_LISTING_BETA] }, { signal })
        for await (const model of listing) {
          const fallback = (model.allowed_fallback_models?.length ?? 0) > 0
          models.push({ id: model.id, label: model.display_name, fallback })
          if (models.length >= MAX_AI_MODELS) break
        }
        return models
      } catch (cause) {
        throw translateSdkError(cause)
      }
    },
    async *stream(prompt, signal): AsyncGenerator<AiStreamPart> {
      try {
        const calls: OpenCalls = { open: new Map(), started: 0 }
        for await (const event of openStream(client, prompt, signal)) {
          const part = partOf(event) ?? toolCallOf(event, calls)
          if (part !== undefined) yield part
        }
      } catch (cause) {
        throw translateSdkError(cause)
      }
    },
  }
}

function openStream(
  client: AnthropicClient,
  prompt: AiPrompt,
  signal: AbortSignal,
): AsyncIterable<BetaRawMessageStreamEvent> {
  return client.beta.messages.stream(
    {
      model: prompt.model,
      max_tokens: prompt.maxTokens,
      system: prompt.system,
      messages: [
        ...prompt.messages.map((turn): BetaMessageParam => ({ role: turn.role, content: turn.text })),
        ...(prompt.exchanges ?? []).flatMap(exchangeMessages),
      ],
      ...toolsOf(prompt),
      ...(prompt.fallback === true ? { betas: [FALLBACK_MESSAGES_BETA], fallbacks: 'default' } : {}),
    },
    { signal },
  )
}

function toolsOf(prompt: AiPrompt): { tools?: BetaTool[] } {
  const tools = prompt.tools ?? []
  if (tools.length === 0) return {}
  return {
    tools: tools.map((tool) => ({
      name: tool.name,
      description: tool.description,
      input_schema: { ...tool.parameters, type: 'object' },
    })),
  }
}

/** The API wants an object; arguments the model got wrong go back as they were, under a key. */
function inputOf(call: AiToolCall): Record<string, unknown> {
  try {
    const input = JSON.parse(call.arguments === '' ? '{}' : call.arguments) as unknown
    if (typeof input === 'object' && input !== null && !Array.isArray(input))
      return input as Record<string, unknown>
  } catch {
    // The tool's result already told the model the JSON was wrong.
  }
  return { arguments: call.arguments }
}

function exchangeMessages(exchange: AiToolExchange): BetaMessageParam[] {
  return [
    {
      role: 'assistant',
      content: [
        ...(exchange.text === '' ? [] : [{ type: 'text' as const, text: exchange.text }]),
        ...exchange.calls.map((call) => ({
          type: 'tool_use' as const,
          id: call.id,
          name: call.name,
          input: inputOf(call),
        })),
      ],
    },
    {
      role: 'user',
      content: exchange.results.map((result) => ({
        type: 'tool_result' as const,
        tool_use_id: result.callId,
        content: result.content,
        is_error: result.isError,
      })),
    },
  ]
}

interface OpenCalls {
  readonly open: Map<number, AiToolCall>
  started: number
}

function startCall(calls: OpenCalls, index: number, id: string, name: string): void {
  calls.started += 1
  if (calls.started > MAX_AI_TOOL_CALLS) throw toolCallsTooLargeError()
  calls.open.set(index, { id, name, arguments: '' })
}

function extendCall(calls: OpenCalls, index: number, piece: string): void {
  const call = calls.open.get(index)
  if (call === undefined) return
  const args = call.arguments + piece
  if (args.length > MAX_AI_TOOL_ARGUMENTS_LENGTH) throw toolCallsTooLargeError()
  calls.open.set(index, { ...call, arguments: args })
}

/** The input arrives as pieces of JSON; the call is whole when its block stops. */
function toolCallOf(event: BetaRawMessageStreamEvent, calls: OpenCalls): AiStreamPart | undefined {
  if (event.type === 'content_block_start' && event.content_block.type === 'tool_use') {
    startCall(calls, event.index, event.content_block.id, event.content_block.name)
  } else if (event.type === 'content_block_delta' && event.delta.type === 'input_json_delta') {
    extendCall(calls, event.index, event.delta.partial_json)
  } else if (event.type === 'content_block_stop') {
    const toolCall = calls.open.get(event.index)
    calls.open.delete(event.index)
    if (toolCall !== undefined) return { toolCall }
  }
  return undefined
}

/** After a fallback the new model continues from the partial text, so nothing already sent is undone. */
function partOf(event: BetaRawMessageStreamEvent): AiStreamPart | undefined {
  if (event.type === 'content_block_delta' && event.delta.type === 'text_delta') {
    return { text: event.delta.text }
  }
  if (event.type === 'content_block_start' && event.content_block.type === 'fallback') {
    return { model: event.content_block.to.model }
  }
  if (event.type === 'message_delta' && event.delta.stop_reason === 'refusal') throw refusedError()
  return undefined
}

function translateSdkError(cause: unknown): unknown {
  if (cause instanceof Anthropic.APIUserAbortError) return cause
  if (cause instanceof Anthropic.AuthenticationError || cause instanceof Anthropic.PermissionDeniedError) {
    return aiError(ErrorCode.AiAuthFailed, 'errors.ai.auth')
  }
  if (cause instanceof Anthropic.RateLimitError)
    return aiError(ErrorCode.AiRateLimited, 'errors.ai.rateLimited')
  if (cause instanceof Anthropic.APIConnectionError) return unreachableError()
  if (cause instanceof Anthropic.APIError) {
    return aiError(
      ErrorCode.AiFailed,
      'errors.ai.failed',
      cause.status === undefined ? undefined : `HTTP ${cause.status}`,
    )
  }
  return cause
}
