import { useEffect, useRef, useState } from 'react'
import type { Editor } from '@tiptap/react'
import { AiProviderKind } from '@services/ai/provider.js'
import { LANGUAGE_NAMES, LANGUAGES, Language, type MessageKey } from '@shared/i18n/index.js'
import { MAX_AI_INSTRUCTION_LENGTH, MAX_AI_SELECTION_LENGTH } from '@shared/limits.js'
import { AiAction, EditCommand } from '@shared/types.js'
import { useT } from '../i18n.js'
import { currentPreferences } from '../state/preferences.js'
import { useWorkspace } from '../state/workspace.js'
import type { AiTargetProblem } from './ai-apply.js'
import { AiModelSelect } from './AiModelSelect.js'
import { authorOf, useAiProviders, type AiProviderInfo } from './useAiProviders.js'
import { useAiRequest, type AiRequest } from './useAiRequest.js'

export interface AiLaunch {
  readonly action: AiAction
  readonly nonce: number
}

export const AI_ACTION_LABELS: Readonly<Record<AiAction, MessageKey>> = {
  [AiAction.Rewrite]: 'ai.action.rewrite',
  [AiAction.Summarize]: 'ai.action.summarize',
  [AiAction.Proofread]: 'ai.action.proofread',
  [AiAction.Translate]: 'ai.action.translate',
  [AiAction.Custom]: 'ai.action.custom',
}

const PROBLEM_LABELS: Readonly<Record<AiTargetProblem, MessageKey>> = {
  empty: 'ai.problem.empty',
  unsupported: 'ai.problem.unsupported',
  tooLong: 'ai.problem.tooLong',
}

interface AiPanelProps {
  readonly editor: Editor
  readonly launch: AiLaunch | null
  readonly onClose: () => void
  readonly onConfigure: () => void
  readonly configuring: boolean
}

interface AiForm {
  readonly action: AiAction
  readonly targetLanguage: Language
  readonly instruction: string
}

function otherLanguage(): Language {
  return currentPreferences().language === Language.English ? Language.Portuguese : Language.English
}

export function AiPanel({
  editor,
  launch,
  onClose,
  onConfigure,
  configuring,
}: AiPanelProps): React.JSX.Element {
  const t = useT()
  const readOnly = useWorkspace((state) => state.readOnly)
  const { list, update } = useAiProviders(configuring)
  const request = useAiRequest(editor)
  const [providerId, setProviderId] = useState<string | null>(null)
  const [form, setForm] = useState<AiForm>(() => ({
    action: AiAction.Rewrite,
    targetLanguage: otherLanguage(),
    instruction: '',
  }))
  const [problem, setProblem] = useState<AiTargetProblem | null>(null)
  const provider =
    list?.providers.find((candidate) => candidate.id === (providerId ?? list.defaultProviderId)) ??
    list?.providers[0]

  const generate = async (chosen: AiForm): Promise<void> => {
    if (provider === undefined) return
    setProblem(
      await request.start({
        providerId: provider.id,
        model: provider.defaultModel ?? '',
        action: chosen.action,
        targetLanguage: chosen.action === AiAction.Translate ? chosen.targetLanguage : undefined,
        instruction: chosen.action === AiAction.Custom ? chosen.instruction : undefined,
      }),
    )
  }
  useLaunch(launch, provider, form, setForm, generate)

  return (
    <aside className="ai-pane" aria-label={t('ai.panel.title')}>
      <div className="nav-pane__header">
        <span className="nav-pane__title">{t('ai.panel.title')}</span>
        <button
          type="button"
          className="nav-pane__close"
          aria-label={t('document.common.close')}
          title={t('document.common.close')}
          onClick={onClose}
        >
          ×
        </button>
      </div>
      {list !== null && provider === undefined ? (
        <AiNoProviders onConfigure={onConfigure} />
      ) : (
        <div className="ai-pane__body">
          {readOnly && <p className="ai-pane__note">{t('ai.panel.readOnly')}</p>}
          <AiRequestForm
            providers={list?.providers ?? []}
            provider={provider}
            onProvider={setProviderId}
            onModel={(model) => {
              if (provider !== undefined)
                void update(window.api.ai.chooseModel({ providerId: provider.id, model }))
            }}
            form={form}
            onForm={setForm}
            running={request.state.phase === 'running'}
            disabled={readOnly || provider === undefined}
            onGenerate={() => void generate(form)}
            onStop={request.cancel}
          />
          {problem !== null && (
            <p className="ai-pane__warning" role="alert">
              {t(PROBLEM_LABELS[problem], { max: MAX_AI_SELECTION_LENGTH })}
            </p>
          )}
          {provider !== undefined && <AiResult request={request} provider={provider} readOnly={readOnly} />}
        </div>
      )}
    </aside>
  )
}

function AiNoProviders({ onConfigure }: { readonly onConfigure: () => void }): React.JSX.Element {
  const t = useT()
  return (
    <div className="ai-pane__body">
      <p className="ai-pane__note">{t('ai.panel.noProviders')}</p>
      <button type="button" className="btn btn--primary" onClick={onConfigure}>
        {t('ai.panel.configure')}
      </button>
    </div>
  )
}

function useLaunch(
  launch: AiLaunch | null,
  provider: AiProviderInfo | undefined,
  form: AiForm,
  setForm: (form: AiForm) => void,
  generate: (form: AiForm) => Promise<void>,
): void {
  const handled = useRef<number | null>(null)
  useEffect(() => {
    if (launch === null || provider === undefined || handled.current === launch.nonce) return
    handled.current = launch.nonce
    const next = { ...form, action: launch.action }
    setForm(next)
    if (launch.action !== AiAction.Custom) void generate(next)
  }, [launch, provider, form, setForm, generate])
}

interface AiRequestFormProps {
  readonly providers: readonly AiProviderInfo[]
  readonly provider: AiProviderInfo | undefined
  readonly onProvider: (id: string) => void
  readonly onModel: (model: string) => void
  readonly form: AiForm
  readonly onForm: (form: AiForm) => void
  readonly running: boolean
  readonly disabled: boolean
  readonly onGenerate: () => void
  readonly onStop: () => void
}

function AiRequestForm(props: AiRequestFormProps): React.JSX.Element {
  const t = useT()
  const { form, onForm } = props
  return (
    <>
      <label className="popover__field">
        <span>{t('ai.panel.provider')}</span>
        <select value={props.provider?.id ?? ''} onChange={(event) => props.onProvider(event.target.value)}>
          {props.providers.map((provider) => (
            <option key={provider.id} value={provider.id}>
              {provider.label}
            </option>
          ))}
        </select>
      </label>
      {props.provider !== undefined && props.provider.kind !== AiProviderKind.ClaudeCli && (
        <AiModelSelect provider={props.provider} disabled={props.running} onChoose={props.onModel} />
      )}
      <label className="popover__field">
        <span>{t('ai.panel.action')}</span>
        <select
          value={form.action}
          onChange={(event) => onForm({ ...form, action: event.target.value as AiAction })}
        >
          {Object.values(AiAction).map((action) => (
            <option key={action} value={action}>
              {t(AI_ACTION_LABELS[action])}
            </option>
          ))}
        </select>
      </label>
      {form.action === AiAction.Translate && (
        <label className="popover__field">
          <span>{t('ai.panel.targetLanguage')}</span>
          <select
            value={form.targetLanguage}
            onChange={(event) => onForm({ ...form, targetLanguage: event.target.value as Language })}
          >
            {LANGUAGES.map((language) => (
              <option key={language} value={language}>
                {LANGUAGE_NAMES[language]}
              </option>
            ))}
          </select>
        </label>
      )}
      {form.action === AiAction.Custom && (
        <label className="popover__field">
          <span>{t('ai.panel.instruction')}</span>
          <textarea
            rows={3}
            value={form.instruction}
            maxLength={MAX_AI_INSTRUCTION_LENGTH}
            placeholder={t('ai.panel.instructionPlaceholder')}
            onChange={(event) => onForm({ ...form, instruction: event.target.value })}
          />
        </label>
      )}
      <div className="ai-pane__actions">
        {props.running ? (
          <button type="button" className="btn" onClick={props.onStop}>
            {t('ai.panel.stop')}
          </button>
        ) : (
          <button
            type="button"
            className="btn btn--primary"
            disabled={props.disabled}
            onMouseDown={(event) => event.preventDefault()}
            onClick={props.onGenerate}
          >
            {t('ai.panel.generate')}
          </button>
        )}
      </div>
    </>
  )
}

const STATUS_LABELS: Partial<Record<AiRequest['state']['phase'], MessageKey>> = {
  running: 'ai.panel.running',
  canceled: 'ai.panel.canceled',
  declined: 'ai.panel.declined',
}

function AiResult({
  request,
  provider,
  readOnly,
}: {
  readonly request: AiRequest
  readonly provider: AiProviderInfo
  readonly readOnly: boolean
}): React.JSX.Element | null {
  const t = useT()
  const output = useRef<HTMLTextAreaElement>(null)
  const { state } = request
  if (state.phase === 'idle') return null
  const status = STATUS_LABELS[state.phase]
  const author = authorOf(provider, state.model)
  const finished = state.phase === 'done' && state.text !== ''
  const copy = (): void => {
    output.current?.select()
    void window.api.edit.run({ command: EditCommand.Copy })
  }
  return (
    <>
      <label className="popover__field">
        <span>{t('ai.panel.result')}</span>
        <textarea ref={output} className="ai-pane__output" readOnly rows={10} value={state.text} />
      </label>
      {status !== undefined && (
        <p className="ai-pane__note" role="status">
          {t(status)}
        </p>
      )}
      {state.costUsd !== undefined && (
        <p className="ai-pane__note">{t('ai.panel.cost', { cost: state.costUsd.toFixed(4) })}</p>
      )}
      {state.changed && <p className="ai-pane__warning">{t('ai.panel.changed')}</p>}
      <div className="ai-pane__actions">
        <button
          type="button"
          className="btn btn--primary"
          disabled={!finished || state.changed || readOnly}
          onClick={() => request.apply(author)}
        >
          {t('ai.panel.apply')}
        </button>
        <button
          type="button"
          className="btn"
          disabled={!finished || readOnly}
          onClick={() => request.insertBelow(author)}
        >
          {t('ai.panel.insertBelow')}
        </button>
        <button type="button" className="btn" disabled={state.text === ''} onClick={copy}>
          {t('ai.panel.copy')}
        </button>
      </div>
      <p className="popover__hint">{t('ai.panel.applyHint')}</p>
    </>
  )
}
