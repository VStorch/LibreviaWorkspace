import { useEffect, useState } from 'react'
import { AiProviderKind } from '@services/ai/provider.js'
import type { IpcRequest, IpcResponse } from '@shared/ipc.js'
import type { IpcChannel } from '@shared/ipc-channels.js'
import type { MessageKey } from '@shared/i18n/index.js'
import {
  DEFAULT_LOCAL_AI_CONTEXT_LENGTH,
  MAX_AI_AGENT_ARGS_LENGTH,
  MAX_AI_DOCUMENT_LENGTH,
  MAX_AI_KEY_LENGTH,
  MAX_AI_URL_LENGTH,
  MAX_NAME_LENGTH,
} from '@shared/limits.js'
import { useT } from '../i18n.js'
import { useWorkspace } from '../state/workspace.js'
import { useAiProviders, type AiProviderInfo, type AiProviders } from './useAiProviders.js'

type ProviderDraft = IpcRequest<typeof IpcChannel.AiProviderSave>
type SaveProvider = (draft: ProviderDraft) => Promise<boolean>
type AiPreferences = IpcRequest<typeof IpcChannel.AiPreferencesSave>
type Candidate = IpcResponse<typeof IpcChannel.AiDiscover>['candidates'][number]

const KIND_LABELS: Readonly<Record<AiProviderKind, MessageKey>> = {
  [AiProviderKind.OpenAiCompatible]: 'ai.kind.openaiCompatible',
  [AiProviderKind.Anthropic]: 'ai.kind.anthropic',
  [AiProviderKind.ClaudeCli]: 'ai.kind.claudeCli',
  [AiProviderKind.Acp]: 'ai.kind.acp',
}

function draftOf(provider: AiProviderInfo, change: Partial<ProviderDraft>): ProviderDraft {
  return {
    id: provider.id,
    kind: provider.kind,
    label: provider.label,
    ...(provider.baseUrl === undefined ? {} : { baseUrl: provider.baseUrl }),
    ...(provider.contextLength === undefined ? {} : { contextLength: provider.contextLength }),
    ...(provider.agentId === undefined ? {} : { agentId: provider.agentId }),
    ...(provider.args === undefined ? {} : { args: provider.args }),
    makeDefault: false,
    revokeConsent: false,
    ...change,
  }
}

export function AiProvidersDialog({ onClose }: { readonly onClose: () => void }): React.JSX.Element {
  const t = useT()
  const providers = useAiProviders()
  const save: SaveProvider = (draft) => providers.update(window.api.ai.saveProvider(draft))
  const chooseProgram: ChooseProgram = (draft) => chooseAgentProgram(providers, draft)
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
            chooseProgram={chooseProgram}
            providers={providers}
          />
        ))}
      </ul>
      <DiscoverSection save={save} />
      <ManualProviderForm save={save} chooseProgram={chooseProgram} />
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
  chooseProgram,
  providers,
}: {
  readonly provider: AiProviderInfo
  readonly isDefault: boolean
  readonly save: SaveProvider
  readonly chooseProgram: ChooseProgram
  readonly providers: AiProviders
}): React.JSX.Element {
  const t = useT()
  const [editing, setEditing] = useState(false)
  if (editing) {
    return (
      <li className="ai-providers__item">
        <ManualProviderForm
          save={save}
          chooseProgram={chooseProgram}
          editing={provider}
          onDone={() => setEditing(false)}
        />
      </li>
    )
  }
  return (
    <li className="ai-providers__item">
      <div className="ai-providers__heading">
        <strong>{provider.label}</strong>
        <span className="popover__hint">
          {t(provider.local ? 'ai.providers.local' : 'ai.providers.remote')}
          {provider.commandName === undefined ? '' : ` · ${provider.commandName}`}
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
        <div
          key={`${candidate.kind}-${candidate.baseUrl ?? candidate.agentId ?? ''}`}
          className="ai-providers__candidate"
        >
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
  readonly args: string
}

function formOf(provider: AiProviderInfo | undefined): ProviderForm {
  return {
    kind: provider?.kind ?? AiProviderKind.OpenAiCompatible,
    label: provider?.label ?? '',
    baseUrl: provider?.baseUrl ?? '',
    apiKey: '',
    contextLength: provider?.contextLength?.toString() ?? '',
    args: provider?.args ?? '',
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
function draftFromForm(form: ProviderForm, defaultLabel: string, editing?: AiProviderInfo): ProviderDraft {
  const label = form.label.trim()
  const apiKey = form.apiKey.trim()
  const isAgent = form.kind === AiProviderKind.Acp
  return {
    ...(editing === undefined ? {} : { id: editing.id }),
    ...(isAgent && editing?.agentId !== undefined ? { agentId: editing.agentId } : {}),
    ...(isAgent ? { args: form.args.trim() } : {}),
    kind: form.kind,
    label: label === '' ? defaultLabel : label,
    ...(form.kind === AiProviderKind.OpenAiCompatible ? { baseUrl: form.baseUrl.trim() } : {}),
    ...(needsKey(form.kind) && apiKey !== '' ? { apiKey } : {}),
    ...contextLengthOf(form),
    makeDefault: false,
    revokeConsent: false,
  }
}

function needsKey(kind: AiProviderKind): boolean {
  return kind === AiProviderKind.OpenAiCompatible || kind === AiProviderKind.Anthropic
}

type ChooseProgram = (draft: ProviderDraft) => Promise<boolean>

/** True only when a program was chosen: canceling the file dialog keeps the form as it was. */
async function chooseAgentProgram(providers: AiProviders, draft: ProviderDraft): Promise<boolean> {
  const call = window.api.ai.chooseAgentExecutable({
    ...(draft.id === undefined ? {} : { id: draft.id }),
    label: draft.label,
    args: draft.args ?? '',
  })
  const result = await call
  return (await providers.update(call)) && result.ok && result.data.chosen
}

function ManualProviderForm({
  save,
  chooseProgram,
  editing,
  onDone,
}: {
  readonly save: SaveProvider
  readonly chooseProgram: ChooseProgram
  readonly editing?: AiProviderInfo
  readonly onDone?: () => void
}): React.JSX.Element {
  const t = useT()
  const [form, setForm] = useState<ProviderForm>(() => formOf(editing))
  const change = (field: Partial<ProviderForm>): void => setForm((previous) => ({ ...previous, ...field }))
  const draft = (): ProviderDraft => draftFromForm(form, t(KIND_LABELS[form.kind]), editing)
  const finish = (done: boolean): void => {
    if (!done) return
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
      <FormActions
        form={form}
        editing={editing}
        onSave={() => void save(draft()).then(finish)}
        onChooseProgram={() => void chooseProgram(draft()).then(finish)}
        {...(onDone === undefined ? {} : { onCancel: onDone })}
      />
    </fieldset>
  )
}

/** A new custom agent starts by choosing its program; a known one, or one already chosen, just saves. */
function FormActions({
  form,
  editing,
  onSave,
  onChooseProgram,
  onCancel,
}: {
  readonly form: ProviderForm
  readonly editing: AiProviderInfo | undefined
  readonly onSave: () => void
  readonly onChooseProgram: () => void
  readonly onCancel?: () => void
}): React.JSX.Element {
  const t = useT()
  const isAgent = form.kind === AiProviderKind.Acp
  const customAgent = isAgent && editing?.agentId === undefined
  const canSave = isAgent
    ? editing !== undefined
    : form.kind !== AiProviderKind.OpenAiCompatible || form.baseUrl.trim() !== ''
  return (
    <div className="ai-pane__actions">
      {canSave && (
        <button type="button" className="btn btn--primary" onClick={onSave}>
          {t('ai.providers.save')}
        </button>
      )}
      {customAgent && (
        <button
          type="button"
          className={editing === undefined ? 'btn btn--primary' : 'btn'}
          onClick={onChooseProgram}
        >
          {t(editing === undefined ? 'ai.providers.chooseProgram' : 'ai.providers.chooseOtherProgram')}
        </button>
      )}
      {onCancel !== undefined && (
        <button type="button" className="btn" onClick={onCancel}>
          {t('document.common.cancel')}
        </button>
      )}
    </div>
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
  const isAgent = form.kind === AiProviderKind.Acp
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
      {isAgent && (
        <TextField
          label={t('ai.providers.args')}
          value={form.args}
          maxLength={MAX_AI_AGENT_ARGS_LENGTH}
          hint={t('ai.providers.argsHint')}
          onChange={(args) => change({ args })}
        />
      )}
      {needsKey(form.kind) && (
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
      <label className="popover__check">
        <input
          type="checkbox"
          checked={preferences.mcpServer}
          onChange={(event) => toggle({ mcpServer: event.target.checked })}
        />
        {t('ai.preferences.mcpServer')}
      </label>
      {preferences.mcpServer && <McpSetup />}
    </fieldset>
  )
}

type McpSetupFormat = 'command' | 'json'

/** What the connected program receives, and how to register the app in it. */
function McpSetup(): React.JSX.Element {
  const t = useT()
  const [setup, setSetup] = useState<Record<McpSetupFormat, string>>({ command: '', json: '' })
  useEffect(() => {
    void window.api.mcp.status({}).then((result) => {
      if (result.ok) setSetup({ command: result.data.clientCommand, json: result.data.clientConfig })
    })
  }, [])
  return (
    <div data-testid="mcp-setup">
      <p className="popover__hint">{t('ai.preferences.mcpWarning')}</p>
      <McpSetupBlock format="command" label={t('ai.preferences.mcpCommand')} text={setup.command} />
      <McpSetupBlock format="json" label={t('ai.preferences.mcpConfig')} text={setup.json} />
    </div>
  )
}

function McpSetupBlock({
  format,
  label,
  text,
}: {
  readonly format: McpSetupFormat
  readonly label: string
  readonly text: string
}): React.JSX.Element {
  const t = useT()
  const [copied, setCopied] = useState(false)
  const copy = async (): Promise<void> => {
    const result = await window.api.mcp.copyConfig({ format })
    setCopied(result.ok)
  }
  return (
    <div data-testid={`mcp-setup-${format}`}>
      <span className="popover__hint">{label}</span>
      <pre className="popover__code">{text}</pre>
      <button type="button" className="btn" onClick={() => void copy()}>
        {t(copied ? 'ai.preferences.mcpCopied' : 'ai.preferences.mcpCopy')}
      </button>
    </div>
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
