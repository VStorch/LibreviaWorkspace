import { useState } from 'react'
import type { Editor } from '@tiptap/react'
import { AiProviderKind } from '@services/ai/provider.js'
import type { IpcRequest, IpcResponse } from '@shared/ipc.js'
import type { IpcChannel } from '@shared/ipc-channels.js'
import type { MessageKey } from '@shared/i18n/index.js'
import { MAX_AI_KEY_LENGTH, MAX_AI_URL_LENGTH, MAX_NAME_LENGTH } from '@shared/limits.js'
import { useT } from '../i18n.js'
import { useWorkspace } from '../state/workspace.js'
import { useAiProviders, type AiProviderInfo, type AiProviders } from './useAiProviders.js'

type ProviderDraft = IpcRequest<typeof IpcChannel.AiProviderSave>
type Candidate = IpcResponse<typeof IpcChannel.AiDiscover>['candidates'][number]
type Model = IpcResponse<typeof IpcChannel.AiModelsList>['models'][number]

const KIND_LABELS: Readonly<Record<AiProviderKind, MessageKey>> = {
  [AiProviderKind.OpenAiCompatible]: 'ai.kind.openaiCompatible',
  [AiProviderKind.Anthropic]: 'ai.kind.anthropic',
  [AiProviderKind.ClaudeCli]: 'ai.kind.claudeCli',
}

function draftOf(provider: AiProviderInfo, change: Partial<ProviderDraft>): ProviderDraft {
  return {
    id: provider.id,
    kind: provider.kind,
    label: provider.label,
    ...(provider.baseUrl === undefined ? {} : { baseUrl: provider.baseUrl }),
    makeDefault: false,
    revokeConsent: false,
    ...change,
  }
}

export function AiProvidersDialog({
  onClose,
}: {
  readonly editor: Editor
  readonly onClose: () => void
}): React.JSX.Element {
  const t = useT()
  const providers = useAiProviders()
  const save = (draft: ProviderDraft): Promise<boolean> => providers.update(window.api.ai.saveProvider(draft))
  const list = providers.list
  return (
    <div
      className="popover popover--wide ai-providers"
      role="dialog"
      aria-label={t('ai.providers.title')}
      onKeyDown={(event) => {
        if (event.key === 'Escape') onClose()
      }}
    >
      <strong className="ai-providers__title">{t('ai.providers.title')}</strong>
      <p className="popover__hint">{t('ai.providers.intro')}</p>
      {list !== null && list.providers.length === 0 && (
        <p className="popover__hint">{t('ai.providers.empty')}</p>
      )}
      <ul className="ai-providers__list">
        {list?.providers.map((provider) => (
          <ProviderRow
            key={provider.id}
            provider={provider}
            isDefault={provider.id === list.defaultProviderId}
            save={save}
            providers={providers}
          />
        ))}
      </ul>
      <DiscoverSection save={save} />
      <ManualProviderForm save={save} />
      <div className="popover__actions">
        <span className="popover__spacer" />
        <button type="button" className="btn" onClick={onClose}>
          {t('document.common.close')}
        </button>
      </div>
    </div>
  )
}

function ProviderRow({
  provider,
  isDefault,
  save,
  providers,
}: {
  readonly provider: AiProviderInfo
  readonly isDefault: boolean
  readonly save: (draft: ProviderDraft) => Promise<boolean>
  readonly providers: AiProviders
}): React.JSX.Element {
  const t = useT()
  return (
    <li className="ai-providers__item">
      <div className="ai-providers__heading">
        <strong>{provider.label}</strong>
        <span className="popover__hint">
          {t(provider.local ? 'ai.providers.local' : 'ai.providers.remote')}
          {isDefault ? ` · ${t('ai.providers.default')}` : ''}
        </span>
      </div>
      {provider.kind !== AiProviderKind.ClaudeCli && <ModelChoice provider={provider} save={save} />}
      <div className="ai-pane__actions">
        {!isDefault && (
          <button
            type="button"
            className="btn"
            onClick={() => void save(draftOf(provider, { makeDefault: true }))}
          >
            {t('ai.providers.makeDefault')}
          </button>
        )}
        {provider.consented && (
          <button
            type="button"
            className="btn"
            onClick={() => void save(draftOf(provider, { revokeConsent: true }))}
          >
            {t('ai.providers.revokeConsent')}
          </button>
        )}
        <button
          type="button"
          className="btn"
          onClick={() => void providers.update(window.api.ai.removeProvider({ providerId: provider.id }))}
        >
          {t('ai.providers.remove')}
        </button>
      </div>
    </li>
  )
}

function ModelChoice({
  provider,
  save,
}: {
  readonly provider: AiProviderInfo
  readonly save: (draft: ProviderDraft) => Promise<boolean>
}): React.JSX.Element {
  const t = useT()
  const showError = useWorkspace((state) => state.showError)
  const [models, setModels] = useState<readonly Model[] | null>(null)
  const loadModels = async (): Promise<void> => {
    const result = await window.api.ai.listModels({ providerId: provider.id })
    if (result.ok) setModels(result.data.models)
    else showError(result.error)
  }
  const chosen = provider.defaultModel ?? ''
  const options = models ?? (chosen === '' ? [] : [{ id: chosen, label: chosen }])
  return (
    <div className="popover__row">
      <label className="popover__field">
        <span>{t('ai.providers.model')}</span>
        <select
          value={chosen}
          onChange={(event) => void save(draftOf(provider, { defaultModel: event.target.value }))}
        >
          <option value="">{t('ai.providers.modelAuto')}</option>
          {options.map((model) => (
            <option key={model.id} value={model.id}>
              {model.label}
            </option>
          ))}
        </select>
      </label>
      <button type="button" className="btn" onClick={() => void loadModels()}>
        {t('ai.providers.loadModels')}
      </button>
    </div>
  )
}

function DiscoverSection({
  save,
}: {
  readonly save: (draft: ProviderDraft) => Promise<boolean>
}): React.JSX.Element {
  const t = useT()
  const showError = useWorkspace((state) => state.showError)
  const [candidates, setCandidates] = useState<readonly Candidate[] | null>(null)
  const [searching, setSearching] = useState(false)
  const discover = async (): Promise<void> => {
    setSearching(true)
    const result = await window.api.ai.discover({})
    setSearching(false)
    if (result.ok) setCandidates(result.data.candidates)
    else showError(result.error)
  }
  return (
    <div className="ai-providers__section">
      <button type="button" className="btn" disabled={searching} onClick={() => void discover()}>
        {t(searching ? 'ai.providers.discovering' : 'ai.providers.discover')}
      </button>
      {candidates !== null && candidates.length === 0 && (
        <p className="popover__hint">{t('ai.providers.nothingFound')}</p>
      )}
      {candidates?.map((candidate) => (
        <div key={`${candidate.kind}-${candidate.baseUrl ?? ''}`} className="ai-providers__candidate">
          <span>
            {candidate.label}
            {candidate.baseUrl === undefined ? '' : ` · ${candidate.baseUrl}`}
          </span>
          <button
            type="button"
            className="btn"
            onClick={() => void save({ ...candidate, makeDefault: false, revokeConsent: false })}
          >
            {t('ai.providers.add')}
          </button>
        </div>
      ))}
    </div>
  )
}

function ManualProviderForm({
  save,
}: {
  readonly save: (draft: ProviderDraft) => Promise<boolean>
}): React.JSX.Element {
  const t = useT()
  const [kind, setKind] = useState<AiProviderKind>(AiProviderKind.OpenAiCompatible)
  const [label, setLabel] = useState('')
  const [baseUrl, setBaseUrl] = useState('')
  const [apiKey, setApiKey] = useState('')
  const needsUrl = kind === AiProviderKind.OpenAiCompatible
  const needsKey = kind !== AiProviderKind.ClaudeCli
  const submit = async (): Promise<void> => {
    const draft: ProviderDraft = {
      kind,
      label: label.trim() === '' ? t(KIND_LABELS[kind]) : label.trim(),
      ...(needsUrl ? { baseUrl: baseUrl.trim() } : {}),
      ...(needsKey && apiKey.trim() !== '' ? { apiKey: apiKey.trim() } : {}),
      makeDefault: false,
      revokeConsent: false,
    }
    if (await save(draft)) {
      setLabel('')
      setBaseUrl('')
      setApiKey('')
    }
  }
  return (
    <fieldset className="popover__fieldset">
      <legend>{t('ai.providers.manual')}</legend>
      <label className="popover__field">
        <span>{t('ai.providers.kind')}</span>
        <select value={kind} onChange={(event) => setKind(event.target.value as AiProviderKind)}>
          {Object.values(AiProviderKind).map((value) => (
            <option key={value} value={value}>
              {t(KIND_LABELS[value])}
            </option>
          ))}
        </select>
      </label>
      <TextField
        label={t('ai.providers.name')}
        value={label}
        maxLength={MAX_NAME_LENGTH}
        onChange={setLabel}
      />
      {needsUrl && (
        <TextField
          label={t('ai.providers.url')}
          type="url"
          value={baseUrl}
          maxLength={MAX_AI_URL_LENGTH}
          placeholder="http://localhost:11434/v1"
          onChange={setBaseUrl}
        />
      )}
      {needsKey && (
        <TextField
          label={t('ai.providers.key')}
          type="password"
          value={apiKey}
          maxLength={MAX_AI_KEY_LENGTH}
          hint={t('ai.providers.keyHint')}
          onChange={setApiKey}
        />
      )}
      <div className="ai-pane__actions">
        <button
          type="button"
          className="btn btn--primary"
          disabled={needsUrl && baseUrl.trim() === ''}
          onClick={() => void submit()}
        >
          {t('ai.providers.save')}
        </button>
      </div>
    </fieldset>
  )
}

function TextField({
  label,
  hint,
  onChange,
  ...input
}: {
  readonly label: string
  readonly value: string
  readonly maxLength: number
  readonly onChange: (value: string) => void
  readonly type?: 'text' | 'url' | 'password'
  readonly placeholder?: string
  readonly hint?: string
}): React.JSX.Element {
  return (
    <label className="popover__field">
      <span>{label}</span>
      <input type="text" autoComplete="off" {...input} onChange={(event) => onChange(event.target.value)} />
      {hint !== undefined && <span className="popover__hint">{hint}</span>}
    </label>
  )
}
