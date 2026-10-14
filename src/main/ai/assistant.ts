import { AI_MAX_OUTPUT_TOKENS, buildMessages } from '@services/ai/prompts.js'
import { AiProviderKind, type AiModel, type AiProvider } from '@services/ai/provider.js'
import { ErrorCode } from '@shared/errors.js'
import type { IpcRequest, IpcResponse } from '@shared/ipc.js'
import type { IpcChannel } from '@shared/ipc-channels.js'
import { MAX_AI_RESPONSE_LENGTH } from '@shared/limits.js'
import { createAnthropicProvider } from './anthropic.js'
import { createDeltaBatcher } from './batcher.js'
import { createClaudeCliProvider } from './claude-cli.js'
import { CLAUDE_COMMAND, findExecutable } from './discovery.js'
import { aiError } from './errors.js'
import { createOpenAiCompatibleProvider } from './openai-compatible.js'
import type { AiModelCache } from './model-cache.js'
import { isLocalProvider, type AiSettings, type StoredAiProvider } from './settings.js'

export type AiGenerateRequest = IpcRequest<typeof IpcChannel.AiGenerate>
export type AiGenerateResponse = IpcResponse<typeof IpcChannel.AiGenerate>

const STREAM_BATCH_MS = 50

export interface GenerateContext {
  readonly settings: AiSettings
  readonly signal: AbortSignal
  readonly push: (delta: string) => void
  readonly confirmConsent: (provider: StoredAiProvider) => Promise<boolean>
  readonly providerFor?: (provider: StoredAiProvider) => Promise<AiProvider>
  readonly models?: AiModelCache
}

interface Relayed {
  readonly costUsd?: number
  readonly model?: string
}

export async function providerFor(settings: AiSettings, stored: StoredAiProvider): Promise<AiProvider> {
  if (stored.kind === AiProviderKind.OpenAiCompatible) {
    const apiKey = settings.keyOf(stored)
    return createOpenAiCompatibleProvider({
      baseUrl: stored.baseUrl ?? '',
      ...(apiKey === undefined ? {} : { apiKey }),
    })
  }
  if (stored.kind === AiProviderKind.Anthropic) {
    const apiKey = settings.keyOf(stored)
    if (apiKey === undefined) throw aiError(ErrorCode.AiAuthFailed, 'errors.ai.auth')
    return createAnthropicProvider(apiKey)
  }
  const command = await findExecutable(CLAUDE_COMMAND)
  if (command === null) throw aiError(ErrorCode.AiNotConfigured, 'errors.ai.cliMissing')
  return createClaudeCliProvider({ command })
}

export async function ensureConsent(
  settings: AiSettings,
  stored: StoredAiProvider,
  confirmConsent: (provider: StoredAiProvider) => Promise<boolean>,
): Promise<boolean> {
  if (isLocalProvider(stored) || stored.consented) return true
  if (!(await confirmConsent(stored))) return false
  settings.markConsented(stored.id)
  return true
}

async function resolveModel(
  request: AiGenerateRequest,
  stored: StoredAiProvider,
  provider: AiProvider,
  context: GenerateContext,
): Promise<string> {
  const chosen = request.model !== '' ? request.model : (stored.defaultModel ?? '')
  if (chosen !== '' || stored.kind === AiProviderKind.ClaudeCli) return chosen
  const load = (): Promise<AiModel[]> => provider.listModels(context.signal)
  const [first] = await (context.models?.get(stored.id, load) ?? load())
  if (first === undefined) throw aiError(ErrorCode.AiNotConfigured, 'errors.ai.modelMissing')
  return first.id
}

export async function generate(
  request: AiGenerateRequest,
  context: GenerateContext,
): Promise<AiGenerateResponse> {
  const { settings, signal } = context
  const stored = settings.find(request.providerId)
  if (!(await ensureConsent(settings, stored, context.confirmConsent))) return { status: 'declined' }

  const batcher = createDeltaBatcher(context.push, STREAM_BATCH_MS)
  try {
    const provider = await (context.providerFor ?? ((p) => providerFor(settings, p)))(stored)
    const chosen = await resolveModel(request, stored, provider, context)
    const relayed = await relay(
      provider,
      { model: chosen, ...buildMessages(request), maxTokens: AI_MAX_OUTPUT_TOKENS },
      signal,
      batcher.add,
    )
    const model = relayed.model ?? chosen
    return {
      status: 'done',
      ...(model === '' ? {} : { model }),
      ...(relayed.costUsd === undefined ? {} : { costUsd: relayed.costUsd }),
    }
  } catch (cause) {
    if (signal.aborted) return { status: 'canceled' }
    throw cause
  } finally {
    batcher.flush()
  }
}

async function relay(
  provider: AiProvider,
  prompt: Parameters<AiProvider['stream']>[0],
  signal: AbortSignal,
  add: (text: string) => void,
): Promise<Relayed> {
  let length = 0
  let relayed: Relayed = {}
  for await (const part of provider.stream(prompt, signal)) {
    if (!('text' in part)) {
      relayed = { ...relayed, ...part }
      continue
    }
    const text = part.text.slice(0, MAX_AI_RESPONSE_LENGTH - length)
    length += text.length
    if (text !== '') add(text)
    if (length >= MAX_AI_RESPONSE_LENGTH) break
  }
  signal.throwIfAborted()
  return relayed
}
