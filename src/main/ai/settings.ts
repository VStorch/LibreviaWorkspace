import { randomUUID } from 'node:crypto'
import Store from 'electron-store'
import { basename } from 'node:path'
import { z } from 'zod'
import { ACP_AGENT_IDS, acpAgentOf } from '@services/ai/acp-agents.js'
import { joinArgs, splitArgs } from '@services/ai/acp-session.js'
import { isLoopback } from '@services/ai/locality.js'
import { AiProviderKind } from '@services/ai/provider.js'
import { AppError, ErrorCode } from '@shared/errors.js'
import type { IpcRequest, IpcResponse } from '@shared/ipc.js'
import type { IpcChannel } from '@shared/ipc-channels.js'
import { DEFAULT_LOCAL_AI_CONTEXT_LENGTH, MAX_AI_DOCUMENT_LENGTH, MAX_AI_PROVIDERS } from '@shared/limits.js'
import { t } from '../i18n.js'
import { aiError } from './errors.js'
import { createSecretBox, type SecretBox } from './secrets.js'

export type AiProviderDraft = IpcRequest<typeof IpcChannel.AiProviderSave>
export type AiProviderList = IpcResponse<typeof IpcChannel.AiProvidersList>
export type AiPreferences = IpcRequest<typeof IpcChannel.AiPreferencesSave>
export type AiAgentExecutableDraft = IpcRequest<typeof IpcChannel.AiAgentChooseExecutable>

const storedProviderSchema = z.object({
  id: z.string(),
  kind: z.enum(AiProviderKind),
  label: z.string(),
  baseUrl: z.string().optional(),
  sealedKey: z.string().optional(),
  consented: z.boolean(),
  /** What the person allowed: before version 2, the selection only; since, the whole document too. */
  consentVersion: z.number().optional(),
  defaultModel: z.string().optional(),
  contextLength: z.number().optional(),
  /** ACP: a known agent, found on the PATH when used, or the program the person chose. */
  agentId: z.enum(ACP_AGENT_IDS).optional(),
  command: z.string().optional(),
  args: z.array(z.string()).optional(),
})

export const DEFAULT_AI_PREFERENCES: AiPreferences = {
  keepFormatting: true,
  anthropicFallback: false,
  chatProposals: false,
}

const settingsSchema = z.object({
  providers: z.array(storedProviderSchema).max(MAX_AI_PROVIDERS),
  defaultProviderId: z.string().nullable(),
  preferences: z
    .object({
      keepFormatting: z.boolean().default(DEFAULT_AI_PREFERENCES.keepFormatting),
      anthropicFallback: z.boolean().default(DEFAULT_AI_PREFERENCES.anthropicFallback),
      chatProposals: z.boolean().default(DEFAULT_AI_PREFERENCES.chatProposals),
    })
    .default(DEFAULT_AI_PREFERENCES),
})

export type StoredAiProvider = z.infer<typeof storedProviderSchema>
export type AiSettingsData = z.infer<typeof settingsSchema>

const EMPTY: AiSettingsData = { providers: [], defaultProviderId: null, preferences: DEFAULT_AI_PREFERENCES }

export interface AiSettingsStore {
  read(): unknown
  write(data: AiSettingsData): void
}

export const AI_CONSENT_VERSION = 2

export function hasCurrentConsent(provider: Pick<StoredAiProvider, 'consented' | 'consentVersion'>): boolean {
  return provider.consented && provider.consentVersion === AI_CONSENT_VERSION
}

/** Local models read little at a time, and the excess is dropped without a word. */
export function contextLengthOf(provider: StoredAiProvider): number {
  if (provider.contextLength !== undefined) return provider.contextLength
  return isLocalProvider(provider) ? DEFAULT_LOCAL_AI_CONTEXT_LENGTH : MAX_AI_DOCUMENT_LENGTH
}

export function isLocalProvider(provider: Pick<StoredAiProvider, 'kind' | 'baseUrl'>): boolean {
  return (
    provider.kind === AiProviderKind.OpenAiCompatible &&
    provider.baseUrl !== undefined &&
    isLoopback(provider.baseUrl)
  )
}

function baseUrlOf(draft: AiProviderDraft): string | undefined {
  if (draft.kind !== AiProviderKind.OpenAiCompatible) return undefined
  if (draft.baseUrl === undefined) throw invalid()
  return draft.baseUrl
}

function keptConsent(
  draft: AiProviderDraft,
  sameDestination: StoredAiProvider | undefined,
): Pick<StoredAiProvider, 'consented' | 'consentVersion'> {
  const consented =
    sameDestination !== undefined && !draft.revokeConsent && hasCurrentConsent(sameDestination)
  return consented ? { consented, consentVersion: AI_CONSENT_VERSION } : { consented }
}

function contextLengthOfDraft(draft: AiProviderDraft): Pick<StoredAiProvider, 'contextLength'> {
  if (draft.kind !== AiProviderKind.OpenAiCompatible || draft.contextLength === undefined) return {}
  return { contextLength: draft.contextLength }
}

type AgentTarget = Pick<StoredAiProvider, 'agentId' | 'command' | 'args'>

function agentTargetOf(
  draft: AiProviderDraft,
  existing: StoredAiProvider | undefined,
  command: string | undefined,
): AgentTarget {
  if (draft.kind !== AiProviderKind.Acp) return {}
  const typed = draft.args === undefined ? undefined : splitArgs(draft.args)
  if (command !== undefined) return { command, args: typed ?? [] }
  if (draft.agentId !== undefined)
    return { agentId: draft.agentId, args: typed ?? [...acpAgentOf(draft.agentId).args] }
  if (existing?.kind !== AiProviderKind.Acp) throw invalid()
  const kept = existing.agentId === undefined ? { command: existing.command } : { agentId: existing.agentId }
  return { ...kept, args: typed ?? existing.args ?? [] }
}

function sameDestination(existing: StoredAiProvider | undefined, next: StoredAiProvider): boolean {
  return (
    existing !== undefined &&
    existing.kind === next.kind &&
    existing.baseUrl === next.baseUrl &&
    existing.agentId === next.agentId &&
    existing.command === next.command &&
    sameArgs(existing.args, next.args)
  )
}

/** Other arguments can make the same program another agent, or send the text elsewhere. */
function sameArgs(a: readonly string[] | undefined, b: readonly string[] | undefined): boolean {
  const left = a ?? []
  const right = b ?? []
  return left.length === right.length && left.every((arg, index) => arg === right[index])
}

function agentInfoOf(provider: StoredAiProvider): Partial<AiProviderList['providers'][number]> {
  return {
    ...(provider.agentId === undefined ? {} : { agentId: provider.agentId }),
    ...(provider.command === undefined ? {} : { commandName: basename(provider.command) }),
    ...(provider.args === undefined ? {} : { args: joinArgs(provider.args) }),
  }
}

const invalid = (): AppError => new AppError(ErrorCode.InvalidRequest, t('errors.ipc.invalidData'))

/** Apart from the editor preferences: keys and consents are not something a preference reset should touch. */
export class AiSettings {
  constructor(
    private readonly store: AiSettingsStore,
    private readonly secrets: SecretBox,
    private readonly newId: () => string = randomUUID,
  ) {}

  list(): AiProviderList {
    const data = this.read()
    return {
      providers: data.providers.map((provider) => ({
        id: provider.id,
        kind: provider.kind,
        label: provider.label,
        ...(provider.baseUrl === undefined ? {} : { baseUrl: provider.baseUrl }),
        hasKey: provider.sealedKey !== undefined,
        local: isLocalProvider(provider),
        consented: hasCurrentConsent(provider),
        ...(provider.defaultModel === undefined ? {} : { defaultModel: provider.defaultModel }),
        ...(provider.contextLength === undefined ? {} : { contextLength: provider.contextLength }),
        ...agentInfoOf(provider),
      })),
      defaultProviderId: data.defaultProviderId,
      preferences: data.preferences,
    }
  }

  preferences(): AiPreferences {
    return this.read().preferences
  }

  find(id: string): StoredAiProvider {
    const provider = this.read().providers.find((candidate) => candidate.id === id)
    if (provider === undefined) throw aiError(ErrorCode.AiNotConfigured, 'errors.ai.providerMissing')
    return provider
  }

  keyOf(provider: StoredAiProvider): string | undefined {
    return provider.sealedKey === undefined ? undefined : this.secrets.open(provider.sealedKey)
  }

  save(draft: AiProviderDraft, command?: string): AiProviderList {
    const data = this.read()
    const existing = draft.id === undefined ? undefined : this.find(draft.id)
    if (existing === undefined && data.providers.length >= MAX_AI_PROVIDERS) throw invalid()
    const provider = this.merge(draft, existing, command)
    const providers =
      existing === undefined
        ? [...data.providers, provider]
        : data.providers.map((candidate) => (candidate.id === provider.id ? provider : candidate))
    const makeDefault = draft.makeDefault || data.defaultProviderId === null
    this.write({ ...data, providers, defaultProviderId: makeDefault ? provider.id : data.defaultProviderId })
    return this.list()
  }

  /** The path comes from main's own file dialog, never from the renderer. */
  saveAgentExecutable(draft: AiAgentExecutableDraft, command: string): AiProviderList {
    return this.save(
      {
        ...(draft.id === undefined ? {} : { id: draft.id }),
        kind: AiProviderKind.Acp,
        label: draft.label,
        args: draft.args,
        makeDefault: false,
        revokeConsent: false,
      },
      command,
    )
  }

  remove(id: string): AiProviderList {
    const data = this.read()
    const providers = data.providers.filter((provider) => provider.id !== id)
    const defaultProviderId =
      data.defaultProviderId === id ? (providers[0]?.id ?? null) : data.defaultProviderId
    this.write({ ...data, providers, defaultProviderId })
    return this.list()
  }

  /** Apart from `save`: picking a model must not touch the key or the consent. */
  chooseModel(id: string, model: string): AiProviderList {
    this.find(id)
    this.updateProvider(id, (provider) => {
      const chosen: StoredAiProvider = { ...provider }
      if (model === '') delete chosen.defaultModel
      else chosen.defaultModel = model
      return chosen
    })
    return this.list()
  }

  savePreferences(preferences: AiPreferences): AiProviderList {
    this.write({ ...this.read(), preferences })
    return this.list()
  }

  markConsented(id: string): void {
    this.updateProvider(id, (provider) => ({
      ...provider,
      consented: true,
      consentVersion: AI_CONSENT_VERSION,
    }))
  }

  private updateProvider(id: string, change: (provider: StoredAiProvider) => StoredAiProvider): void {
    const data = this.read()
    this.write({
      ...data,
      providers: data.providers.map((provider) => (provider.id === id ? change(provider) : provider)),
    })
  }

  private merge(
    draft: AiProviderDraft,
    existing: StoredAiProvider | undefined,
    command: string | undefined,
  ): StoredAiProvider {
    const baseUrl = baseUrlOf(draft)
    const destination: StoredAiProvider = {
      id: existing?.id ?? this.newId(),
      kind: draft.kind,
      label: draft.label,
      consented: false,
      ...(baseUrl === undefined ? {} : { baseUrl }),
      ...agentTargetOf(draft, existing, command),
    }
    const kept = sameDestination(existing, destination) ? existing : undefined
    const sealedKey = this.sealedKeyOf(draft, kept)
    const defaultModel = draft.defaultModel ?? kept?.defaultModel ?? ''
    return {
      ...destination,
      ...(sealedKey === undefined ? {} : { sealedKey }),
      ...keptConsent(draft, kept),
      ...(defaultModel === '' ? {} : { defaultModel }),
      ...contextLengthOfDraft(draft),
    }
  }

  private sealedKeyOf(
    draft: AiProviderDraft,
    sameDestination: StoredAiProvider | undefined,
  ): string | undefined {
    if (draft.kind === AiProviderKind.ClaudeCli || draft.kind === AiProviderKind.Acp) return undefined
    if (draft.apiKey === undefined) return sameDestination?.sealedKey
    return draft.apiKey === '' ? undefined : this.secrets.seal(draft.apiKey)
  }

  private read(): AiSettingsData {
    const parsed = settingsSchema.safeParse(this.store.read())
    return parsed.success ? parsed.data : EMPTY
  }

  private write(data: AiSettingsData): void {
    this.store.write(data)
  }
}

let instance: AiSettings | null = null

export function aiSettings(): AiSettings {
  if (instance === null) {
    const store = new Store<AiSettingsData>({ name: 'ai', defaults: EMPTY, clearInvalidConfig: true })
    instance = new AiSettings(
      { read: () => store.store, write: (data) => (store.store = data) },
      createSecretBox(),
    )
  }
  return instance
}
