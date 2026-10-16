import { acpAgentOf } from '@services/ai/acp-agents.js'
import { joinArgs } from '@services/ai/acp-session.js'
import { buildChatPrompt, buildToolChatPrompt, type AiChatPrompt } from '@services/ai/chat.js'
import { contextWindow } from '@services/ai/context.js'
import { AI_MAX_OUTPUT_TOKENS, buildMessages } from '@services/ai/prompts.js'
import type { AiProposal } from '@services/ai/proposals.js'
import {
  AiProviderKind,
  AiRole,
  offersDocumentTools,
  type AiModel,
  type AiPrompt,
  type AiProvider,
} from '@services/ai/provider.js'
import { ErrorCode } from '@shared/errors.js'
import type { IpcRequest, IpcResponse } from '@shared/ipc.js'
import type { IpcChannel } from '@shared/ipc-channels.js'
import { createAcpProvider } from './acp.js'
import { createAnthropicProvider } from './anthropic.js'
import { createDeltaBatcher } from './batcher.js'
import { createClaudeCliProvider } from './claude-cli.js'
import { CLAUDE_COMMAND, findExecutable } from './discovery.js'
import { aiError } from './errors.js'
import { createOpenAiCompatibleProvider } from './openai-compatible.js'
import type { AiModelCache } from './model-cache.js'
import { relay, type RelayMeta } from './relay.js'
import {
  contextLengthOf,
  hasCurrentConsent,
  isLocalProvider,
  type AiSettings,
  type StoredAiProvider,
} from './settings.js'
import { relayWithTools } from './tool-loop.js'

export type AiGenerateRequest = IpcRequest<typeof IpcChannel.AiGenerate>
export type AiGenerateResponse = IpcResponse<typeof IpcChannel.AiGenerate>
export type AiConversationRequest = IpcRequest<typeof IpcChannel.AiChat>
export type AiConversationResponse = IpcResponse<typeof IpcChannel.AiChat>

interface PromptTarget {
  readonly providerId: string
  readonly model: string
}

interface Answered extends RelayMeta {
  readonly proposals?: readonly AiProposal[]
}

type Runner = (
  provider: AiProvider,
  prompt: AiPrompt,
  signal: AbortSignal,
  add: (text: string) => void,
) => Promise<Answered>

interface Plan {
  readonly prompt: AiChatPrompt
  readonly run: Runner
}

type Planner = (stored: StoredAiProvider) => Plan

const STREAM_BATCH_MS = 50

export interface GenerateContext {
  readonly settings: AiSettings
  readonly signal: AbortSignal
  readonly push: (delta: string) => void
  readonly confirmConsent: (provider: StoredAiProvider) => Promise<boolean>
  readonly providerFor?: (provider: StoredAiProvider) => Promise<AiProvider>
  readonly models?: AiModelCache
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
  if (stored.kind === AiProviderKind.Acp) return acpProviderFor(stored)
  const command = await findExecutable(CLAUDE_COMMAND)
  if (command === null) throw aiError(ErrorCode.AiNotConfigured, 'errors.ai.cliMissing')
  return createClaudeCliProvider({ command })
}

/** A known agent is looked up on the PATH each time; a chosen one runs from where it was chosen. */
async function acpProviderFor(stored: StoredAiProvider): Promise<AiProvider> {
  const command =
    stored.agentId === undefined
      ? (stored.command ?? null)
      : await findExecutable(acpAgentOf(stored.agentId).command)
  if (command === null) throw aiError(ErrorCode.AiNotConfigured, 'errors.ai.agentMissing')
  return createAcpProvider({ command, args: stored.args ?? [] })
}

/** What the consent names: for an agent, the program that will run and how, not the name given. */
export function consentDestination(stored: StoredAiProvider): string | undefined {
  if (stored.kind !== AiProviderKind.Acp) return stored.baseUrl
  const program = stored.command ?? (stored.agentId === undefined ? '' : acpAgentOf(stored.agentId).command)
  return joinArgs([program, ...(stored.args ?? [])])
}

export async function ensureConsent(
  settings: AiSettings,
  stored: StoredAiProvider,
  confirmConsent: (provider: StoredAiProvider) => Promise<boolean>,
): Promise<boolean> {
  if (isLocalProvider(stored) || hasCurrentConsent(stored)) return true
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
  request: PromptTarget,
  stored: StoredAiProvider,
  provider: AiProvider,
  context: GenerateContext,
): Promise<ResolvedModel> {
  const chosen = request.model !== '' ? request.model : (stored.defaultModel ?? '')
  const wantsFallback =
    stored.kind === AiProviderKind.Anthropic && context.settings.preferences().anthropicFallback
  const automaticOnly = stored.kind === AiProviderKind.ClaudeCli || stored.kind === AiProviderKind.Acp
  if (automaticOnly || (chosen !== '' && !wantsFallback)) {
    return { id: chosen, fallback: false }
  }
  const models = await listedModels(stored, provider, context)
  const model = chosen === '' ? models[0] : models.find((candidate) => candidate.id === chosen)
  if (model === undefined && chosen === '') {
    throw aiError(ErrorCode.AiNotConfigured, 'errors.ai.modelMissing')
  }
  return { id: model?.id ?? chosen, fallback: wantsFallback && model?.fallback === true }
}

function messagesFor(request: AiGenerateRequest, stored: StoredAiProvider): AiChatPrompt {
  const budget = contextLengthOf(stored) - request.text.length
  const { system, user } = buildMessages({
    ...request,
    context: contextWindow(request.before, request.after, budget),
  })
  return { system, messages: [{ role: AiRole.User, text: user }] }
}

const relayOnce: Runner = (provider, prompt, signal, add) => relay(provider, prompt, signal, add)

export function generate(request: AiGenerateRequest, context: GenerateContext): Promise<AiGenerateResponse> {
  return answer(request, context, (stored) => ({ prompt: messagesFor(request, stored), run: relayOnce }))
}

/**
 * The conversation sends the document as it is now, with each message. With the proposals on and
 * a provider that takes tools, the model reads it by paragraph and proposes; the person applies.
 */
export function chat(
  request: AiConversationRequest,
  context: GenerateContext,
): Promise<AiConversationResponse> {
  return answer(request, context, (stored) => {
    const contextLength = contextLengthOf(stored)
    const wantsTools = request.proposals && context.settings.preferences().chatProposals
    if (!wantsTools || !offersDocumentTools(stored.kind)) {
      return { prompt: buildChatPrompt(request, contextLength), run: relayOnce }
    }
    return {
      prompt: buildToolChatPrompt(request, contextLength),
      run: (provider, prompt, signal, add) =>
        relayWithTools(
          provider,
          prompt,
          { paragraphs: request.paragraphs, readBudget: contextLength },
          signal,
          add,
        ),
    }
  })
}

async function answer(
  request: PromptTarget,
  context: GenerateContext,
  plan: Planner,
): Promise<AiConversationResponse> {
  const { settings, signal } = context
  const stored = settings.find(request.providerId)
  if (!(await ensureConsent(settings, stored, context.confirmConsent))) return { status: 'declined' }

  const batcher = createDeltaBatcher(context.push, STREAM_BATCH_MS)
  try {
    const provider = await (context.providerFor ?? ((p) => providerFor(settings, p)))(stored)
    const chosen = await resolveModel(request, stored, provider, context)
    const { prompt, run } = plan(stored)
    const answered = await run(
      provider,
      {
        model: chosen.id,
        ...prompt,
        maxTokens: AI_MAX_OUTPUT_TOKENS,
        ...(chosen.fallback ? { fallback: true } : {}),
      },
      signal,
      batcher.add,
    )
    const model = answered.model ?? chosen.id
    return {
      status: 'done',
      ...(model === '' ? {} : { model }),
      ...(answered.costUsd === undefined ? {} : { costUsd: answered.costUsd }),
      ...(answered.proposals === undefined ? {} : { proposals: [...answered.proposals] }),
    }
  } catch (cause) {
    if (signal.aborted) return { status: 'canceled' }
    throw cause
  } finally {
    batcher.flush()
  }
}
