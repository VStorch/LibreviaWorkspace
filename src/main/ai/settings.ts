import { randomUUID } from 'node:crypto'
import Store from 'electron-store'
import { z } from 'zod'
import { isLoopback } from '@services/ai/locality.js'
import { AiProviderKind } from '@services/ai/provider.js'
import { AppError, ErrorCode } from '@shared/errors.js'
import type { IpcRequest, IpcResponse } from '@shared/ipc.js'
import type { IpcChannel } from '@shared/ipc-channels.js'
import { MAX_AI_PROVIDERS } from '@shared/limits.js'
import { t } from '../i18n.js'
import { aiError } from './errors.js'
import { createSecretBox, type SecretBox } from './secrets.js'

export type AiProviderDraft = IpcRequest<typeof IpcChannel.AiProviderSave>
export type AiProviderList = IpcResponse<typeof IpcChannel.AiProvidersList>

const storedProviderSchema = z.object({
  id: z.string(),
  kind: z.enum(AiProviderKind),
  label: z.string(),
  baseUrl: z.string().optional(),
  sealedKey: z.string().optional(),
  consented: z.boolean(),
  defaultModel: z.string().optional(),
})

const settingsSchema = z.object({
  providers: z.array(storedProviderSchema).max(MAX_AI_PROVIDERS),
  defaultProviderId: z.string().nullable(),
})

export type StoredAiProvider = z.infer<typeof storedProviderSchema>
export type AiSettingsData = z.infer<typeof settingsSchema>

const EMPTY: AiSettingsData = { providers: [], defaultProviderId: null }

export interface AiSettingsStore {
  read(): unknown
  write(data: AiSettingsData): void
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
        consented: provider.consented,
        ...(provider.defaultModel === undefined ? {} : { defaultModel: provider.defaultModel }),
      })),
      defaultProviderId: data.defaultProviderId,
    }
  }

  find(id: string): StoredAiProvider {
    const provider = this.read().providers.find((candidate) => candidate.id === id)
    if (provider === undefined) throw aiError(ErrorCode.AiNotConfigured, 'errors.ai.providerMissing')
    return provider
  }

  keyOf(provider: StoredAiProvider): string | undefined {
    return provider.sealedKey === undefined ? undefined : this.secrets.open(provider.sealedKey)
  }

  save(draft: AiProviderDraft): AiProviderList {
    const data = this.read()
    const existing = draft.id === undefined ? undefined : this.find(draft.id)
    if (existing === undefined && data.providers.length >= MAX_AI_PROVIDERS) throw invalid()
    const provider = this.merge(draft, existing)
    const providers =
      existing === undefined
        ? [...data.providers, provider]
        : data.providers.map((candidate) => (candidate.id === provider.id ? provider : candidate))
    const makeDefault = draft.makeDefault || data.defaultProviderId === null
    this.write({ providers, defaultProviderId: makeDefault ? provider.id : data.defaultProviderId })
    return this.list()
  }

  remove(id: string): AiProviderList {
    const data = this.read()
    const providers = data.providers.filter((provider) => provider.id !== id)
    const defaultProviderId =
      data.defaultProviderId === id ? (providers[0]?.id ?? null) : data.defaultProviderId
    this.write({ providers, defaultProviderId })
    return this.list()
  }

  markConsented(id: string): void {
    const data = this.read()
    this.write({
      ...data,
      providers: data.providers.map((provider) =>
        provider.id === id ? { ...provider, consented: true } : provider,
      ),
    })
  }

  private merge(draft: AiProviderDraft, existing: StoredAiProvider | undefined): StoredAiProvider {
    const baseUrl = baseUrlOf(draft)
    const kept = existing?.kind === draft.kind && existing.baseUrl === baseUrl ? existing : undefined
    const sealedKey = this.sealedKeyOf(draft, kept)
    const defaultModel = draft.defaultModel ?? existing?.defaultModel ?? ''
    return {
      id: existing?.id ?? this.newId(),
      kind: draft.kind,
      label: draft.label,
      ...(baseUrl === undefined ? {} : { baseUrl }),
      ...(sealedKey === undefined ? {} : { sealedKey }),
      consented: kept !== undefined && !draft.revokeConsent && kept.consented,
      ...(defaultModel === '' ? {} : { defaultModel }),
    }
  }

  private sealedKeyOf(
    draft: AiProviderDraft,
    sameDestination: StoredAiProvider | undefined,
  ): string | undefined {
    if (draft.kind === AiProviderKind.ClaudeCli) return undefined
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
