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

interface ResolvedModel {
  readonly id: string
  readonly fallback: boolean
}

function listedModels(
  stored: StoredAiProvider,
  provider: AiProvider,
  context: GenerateContext,
): Promise<readonly AiModel[]> {
  const load = (): Promise<AiModel[]> => provider.listModels(context.signal)
  return context.models?.get(stored.id, load) ?? load()
}

async function resolveModel(
  request: AiGenerateRequest,
  stored: StoredAiProvider,
  provider: AiProvider,
  context: GenerateContext,
): Promise<ResolvedModel> {
  const chosen = request.model !== '' ? request.model : (stored.defaultModel ?? '')
  const wantsFallback =
    stored.kind === AiProviderKind.Anthropic && context.settings.preferences().anthropicFallback
  if (stored.kind === AiProviderKind.ClaudeCli || (chosen !== '' && !wantsFallback)) {
    return { id: chosen, fallback: false }
  }
  const models = await listedModels(stored, provider, context)
  const model = chosen === '' ? models[0] : models.find((candidate) => candidate.id === chosen)
  if (model === undefined && chosen === '') {
    throw aiError(ErrorCode.AiNotConfigured, 'errors.ai.modelMissing')
  }
  return { id: model?.id ?? chosen, fallback: wantsFallback && model?.fallback === true }
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
      {
        model: chosen.id,
        ...buildMessages(request),
        maxTokens: AI_MAX_OUTPUT_TOKENS,
        ...(chosen.fallback ? { fallback: true } : {}),
      },
      signal,
      batcher.add,
    )
    const model = relayed.model ?? chosen.id
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
