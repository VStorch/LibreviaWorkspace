import Anthropic from '@anthropic-ai/sdk'
import type { AiModel, AiProvider, AiStreamPart } from '@services/ai/provider.js'
import { ErrorCode } from '@shared/errors.js'
import { MAX_AI_MODELS } from '@shared/limits.js'
import { aiError, refusedError, unreachableError } from './errors.js'

type AnthropicClient = Pick<Anthropic, 'models' | 'messages'>

export function createAnthropicProvider(
  apiKey: string,
  client: AnthropicClient = new Anthropic({ apiKey }),
): AiProvider {
  return {
    async listModels(signal): Promise<AiModel[]> {
      try {
        const models: AiModel[] = []
        for await (const model of client.models.list({}, { signal })) {
          models.push({ id: model.id, label: model.display_name })
          if (models.length >= MAX_AI_MODELS) break
        }
        return models
      } catch (cause) {
        throw translateSdkError(cause)
      }
    },
    async *stream(prompt, signal): AsyncGenerator<AiStreamPart> {
      try {
        const stream = client.messages.stream(
          {
            model: prompt.model,
            max_tokens: prompt.maxTokens,
            system: prompt.system,
            messages: [{ role: 'user', content: prompt.user }],
          },
          { signal },
        )
        for await (const event of stream) {
          if (event.type === 'content_block_delta' && event.delta.type === 'text_delta') {
            yield { text: event.delta.text }
          } else if (event.type === 'message_delta' && event.delta.stop_reason === 'refusal') {
            throw refusedError()
          }
        }
      } catch (cause) {
        throw translateSdkError(cause)
      }
    },
  }
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
