import { useState } from 'react'
import type { Editor } from '@tiptap/react'
import { AiProviderKind } from '@services/ai/provider.js'
import type { IpcRequest, IpcResponse } from '@shared/ipc.js'
import type { IpcChannel } from '@shared/ipc-channels.js'
import type { MessageKey } from '@shared/i18n/index.js'
import {
  DEFAULT_LOCAL_AI_CONTEXT_LENGTH,
  MAX_AI_DOCUMENT_LENGTH,
  MAX_AI_KEY_LENGTH,
  MAX_AI_URL_LENGTH,
  MAX_NAME_LENGTH,
} from '@shared/limits.js'
import { useT } from '../i18n.js'
import { useWorkspace } from '../state/workspace.js'
import { useAiProviders, type AiProviderInfo, type AiProviders } from './useAiProviders.js'

type ProviderDraft = IpcRequest<typeof IpcChannel.AiProviderSave>
type AiPreferences = IpcRequest<typeof IpcChannel.AiPreferencesSave>
type Candidate = IpcResponse<typeof IpcChannel.AiDiscover>['candidates'][number]

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
    ...(provider.contextLength === undefined ? {} : { contextLength: provider.contextLength }),
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
      <PreferencesSection providers={providers} />
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
  const [editing, setEditing] = useState(false)
  if (editing) {
    return (
      <li className="ai-providers__item">
        <ManualProviderForm save={save} editing={provider} onDone={() => setEditing(false)} />
      </li>
    )
  }
  return (
    <li className="ai-providers__item">
      <div className="ai-providers__heading">
        <strong>{provider.label}</strong>
        <span className="popover__hint">
          {t(provider.local ? 'ai.providers.local' : 'ai.providers.remote')}
          {isDefault ? ` · ${t('ai.providers.default')}` : ''}
        </span>
      </div>
      <div className="ai-pane__actions">
        <button type="button" className="btn" onClick={() => setEditing(true)}>
          {t('ai.providers.edit')}
        </button>
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

interface ProviderForm {
  readonly kind: AiProviderKind
  readonly label: string
  readonly baseUrl: string
  readonly apiKey: string
  readonly contextLength: string
}

function formOf(provider: AiProviderInfo | undefined): ProviderForm {
  return {
    kind: provider?.kind ?? AiProviderKind.OpenAiCompatible,
    label: provider?.label ?? '',
    baseUrl: provider?.baseUrl ?? '',
    apiKey: '',
    contextLength: provider?.contextLength?.toString() ?? '',
  }
}

const CONTEXT_LENGTH_DIGITS = String(MAX_AI_DOCUMENT_LENGTH).length

/** Blank goes back to the default; a number out of range is brought inside it. */
function contextLengthOf(form: ProviderForm): { contextLength?: number } {
  const typed = Number.parseInt(form.contextLength, 10)
  if (form.kind !== AiProviderKind.OpenAiCompatible || Number.isNaN(typed)) return {}
  return { contextLength: Math.min(MAX_AI_DOCUMENT_LENGTH, Math.max(0, typed)) }
}

/** A blank key while editing keeps the stored one; main forgets it anyway if the address changes. */
function draftFromForm(form: ProviderForm, defaultLabel: string, id: string | undefined): ProviderDraft {
  const label = form.label.trim()
  const apiKey = form.apiKey.trim()
  return {
    ...(id === undefined ? {} : { id }),
    kind: form.kind,
    label: label === '' ? defaultLabel : label,
    ...(form.kind === AiProviderKind.OpenAiCompatible ? { baseUrl: form.baseUrl.trim() } : {}),
    ...(form.kind !== AiProviderKind.ClaudeCli && apiKey !== '' ? { apiKey } : {}),
    ...contextLengthOf(form),
    makeDefault: false,
    revokeConsent: false,
  }
}

function ManualProviderForm({
  save,
  editing,
  onDone,
}: {
  readonly save: (draft: ProviderDraft) => Promise<boolean>
  readonly editing?: AiProviderInfo
  readonly onDone?: () => void
}): React.JSX.Element {
  const t = useT()
  const [form, setForm] = useState<ProviderForm>(() => formOf(editing))
  const change = (field: Partial<ProviderForm>): void => setForm((previous) => ({ ...previous, ...field }))
  const needsUrl = form.kind === AiProviderKind.OpenAiCompatible
  const submit = async (): Promise<void> => {
    if (!(await save(draftFromForm(form, t(KIND_LABELS[form.kind]), editing?.id)))) return
    setForm(formOf(undefined))
    onDone?.()
  }
  return (
    <fieldset className="popover__fieldset">
      <legend>{t(editing === undefined ? 'ai.providers.manual' : 'ai.providers.editing')}</legend>
      <label className="popover__field">
        <span>{t('ai.providers.kind')}</span>
        <select
          value={form.kind}
          disabled={editing !== undefined}
          onChange={(event) => change({ kind: event.target.value as AiProviderKind })}
        >
          {Object.values(AiProviderKind).map((value) => (
            <option key={value} value={value}>
              {t(KIND_LABELS[value])}
            </option>
          ))}
        </select>
      </label>
      <TextField
        label={t('ai.providers.name')}
        value={form.label}
        maxLength={MAX_NAME_LENGTH}
        onChange={(label) => change({ label })}
      />
      <ConnectionFields form={form} change={change} keyStored={editing?.hasKey === true} />
      <div className="ai-pane__actions">
        <button
          type="button"
          className="btn btn--primary"
          disabled={needsUrl && form.baseUrl.trim() === ''}
          onClick={() => void submit()}
        >
          {t('ai.providers.save')}
        </button>
        {onDone !== undefined && (
          <button type="button" className="btn" onClick={onDone}>
            {t('document.common.cancel')}
          </button>
        )}
      </div>
    </fieldset>
  )
}

function ConnectionFields({
  form,
  change,
  keyStored,
}: {
  readonly form: ProviderForm
  readonly change: (field: Partial<ProviderForm>) => void
  readonly keyStored: boolean
}): React.JSX.Element {
  const t = useT()
  const needsUrl = form.kind === AiProviderKind.OpenAiCompatible
  const needsKey = form.kind !== AiProviderKind.ClaudeCli
  return (
    <>
      {needsUrl && (
        <TextField
          label={t('ai.providers.url')}
          type="url"
          value={form.baseUrl}
          maxLength={MAX_AI_URL_LENGTH}
          placeholder="http://localhost:11434/v1"
          onChange={(baseUrl) => change({ baseUrl })}
        />
      )}
      {needsUrl && (
        <TextField
          label={t('ai.providers.contextLength')}
          type="number"
          value={form.contextLength}
          maxLength={CONTEXT_LENGTH_DIGITS}
          hint={t('ai.providers.contextLengthHint', {
            local: DEFAULT_LOCAL_AI_CONTEXT_LENGTH,
            remote: MAX_AI_DOCUMENT_LENGTH,
          })}
          onChange={(contextLength) => change({ contextLength })}
        />
      )}
      {needsKey && (
        <TextField
          label={t('ai.providers.key')}
          type="password"
          value={form.apiKey}
          maxLength={MAX_AI_KEY_LENGTH}
          hint={t(keyStored ? 'ai.providers.keyKeep' : 'ai.providers.keyHint')}
          onChange={(apiKey) => change({ apiKey })}
        />
      )}
    </>
  )
}

function PreferencesSection({ providers }: { readonly providers: AiProviders }): React.JSX.Element | null {
  const t = useT()
  const preferences = providers.list?.preferences
  if (preferences === undefined) return null
  const toggle = (change: Partial<AiPreferences>): void =>
    void providers.update(window.api.ai.savePreferences({ ...preferences, ...change }))
  return (
    <fieldset className="popover__fieldset">
      <legend>{t('ai.preferences.title')}</legend>
      <label className="popover__check">
        <input
          type="checkbox"
          checked={preferences.keepFormatting}
          onChange={(event) => toggle({ keepFormatting: event.target.checked })}
        />
        {t('ai.preferences.keepFormatting')}
      </label>
      <label className="popover__check">
        <input
          type="checkbox"
          checked={preferences.anthropicFallback}
          onChange={(event) => toggle({ anthropicFallback: event.target.checked })}
        />
        {t('ai.preferences.anthropicFallback')}
      </label>
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
  readonly type?: 'text' | 'url' | 'password' | 'number'
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
