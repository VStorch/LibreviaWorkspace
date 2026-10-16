import Anthropic from '@anthropic-ai/sdk'
import type { BetaRawMessageStreamEvent } from '@anthropic-ai/sdk/resources/beta/messages/messages.js'
import type { AiModel, AiPrompt, AiProvider, AiStreamPart } from '@services/ai/provider.js'
import { ErrorCode } from '@shared/errors.js'
import { MAX_AI_MODELS } from '@shared/limits.js'
import { aiError, refusedError, unreachableError } from './errors.js'

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
        for await (const event of openStream(client, prompt, signal)) {
          const part = partOf(event)
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
      messages: prompt.messages.map((turn) => ({ role: turn.role, content: turn.text })),
      ...(prompt.fallback === true ? { betas: [FALLBACK_MESSAGES_BETA], fallbacks: 'default' } : {}),
    },
    { signal },
  )
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
