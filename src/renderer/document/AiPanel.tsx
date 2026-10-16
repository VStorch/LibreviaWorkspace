import { useEffect, useRef, useState } from 'react'
import type { Editor } from '@tiptap/react'
import { responseFormat } from '@services/ai/prompts.js'
import { AiProviderKind } from '@services/ai/provider.js'
import { LANGUAGE_NAMES, LANGUAGES, Language, type MessageKey } from '@shared/i18n/index.js'
import { MAX_AI_INSTRUCTION_LENGTH, MAX_AI_SELECTION_LENGTH } from '@shared/limits.js'
import { AiAction, AiFormat } from '@shared/types.js'
import { useT } from '../i18n.js'
import { currentPreferences } from '../state/preferences.js'
import { useWorkspace } from '../state/workspace.js'
import type { AiTargetProblem } from './ai-apply.js'
import { copyContents, copySelection } from './ai-copy.js'
import { AI_STATUS_LABELS, AiChatTab } from './AiChatTab.js'
import { AiMarkdownPreview } from './AiMarkdownPreview.js'
import { AiModelSelect } from './AiModelSelect.js'
import { authorOf, useAiProviders, type AiProviderInfo } from './useAiProviders.js'
import type { AiChat } from './useAiChat.js'
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
  readonly chat: AiChat
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

type AiTab = 'actions' | 'chat'

const AI_TABS: readonly { readonly id: AiTab; readonly label: MessageKey }[] = [
  { id: 'actions', label: 'ai.tab.actions' },
  { id: 'chat', label: 'ai.tab.chat' },
]

/** A launch from the context menu is an action: it shows the tab that runs it. */
function useTab(launch: AiLaunch | null): [AiTab, (tab: AiTab) => void] {
  const [tab, setTab] = useState<AiTab>('actions')
  const [seenLaunch, setSeenLaunch] = useState(launch?.nonce)
  if (launch !== null && launch.nonce !== seenLaunch) {
    setSeenLaunch(launch.nonce)
    setTab('actions')
  }
  return [tab, setTab]
}

export function AiPanel({
  editor,
  chat,
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
  const [tab, setTab] = useTab(launch)
  const provider =
    list?.providers.find((candidate) => candidate.id === (providerId ?? list.defaultProviderId)) ??
    list?.providers[0]

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
          <AiTabs active={tab} onChoose={setTab} />
          {readOnly && <p className="ai-pane__note">{t('ai.panel.readOnly')}</p>}
          <AiProviderFields
            providers={list?.providers ?? []}
            provider={provider}
            onProvider={setProviderId}
            onModel={(model) => {
              if (provider !== undefined)
                void update(window.api.ai.chooseModel({ providerId: provider.id, model }))
            }}
            running={request.state.phase === 'running' || chat.running}
          />
          <AiTabPanel tab="actions" active={tab}>
            <AiActions
              request={request}
              provider={provider}
              keepFormatting={list?.preferences.keepFormatting ?? true}
              readOnly={readOnly}
              launch={launch}
            />
          </AiTabPanel>
          <AiTabPanel tab="chat" active={tab}>
            {provider !== undefined && <AiChatTab chat={chat} provider={provider} readOnly={readOnly} />}
          </AiTabPanel>
        </div>
      )}
    </aside>
  )
}

const tabId = (tab: AiTab): string => `ai-tab-${tab}`
const tabPanelId = (tab: AiTab): string => `ai-tabpanel-${tab}`

function nextTabIndex(key: string, index: number): number | null {
  if (key === 'ArrowRight') return (index + 1) % AI_TABS.length
  if (key === 'ArrowLeft') return (index - 1 + AI_TABS.length) % AI_TABS.length
  if (key === 'Home') return 0
  if (key === 'End') return AI_TABS.length - 1
  return null
}

function AiTabs({
  active,
  onChoose,
}: {
  readonly active: AiTab
  readonly onChoose: (tab: AiTab) => void
}): React.JSX.Element {
  const t = useT()
  const buttons = useRef(new Map<AiTab, HTMLButtonElement>())
  const move = (event: React.KeyboardEvent): void => {
    const next = nextTabIndex(
      event.key,
      AI_TABS.findIndex((tab) => tab.id === active),
    )
    if (next === null) return
    event.preventDefault()
    const { id } = AI_TABS[next]!
    onChoose(id)
    buttons.current.get(id)?.focus()
  }
  return (
    <div className="ai-tabs" role="tablist" aria-label={t('ai.panel.title')} onKeyDown={move}>
      {AI_TABS.map((tab) => (
        <button
          key={tab.id}
          ref={(button) => {
            if (button === null) buttons.current.delete(tab.id)
            else buttons.current.set(tab.id, button)
          }}
          type="button"
          role="tab"
          id={tabId(tab.id)}
          aria-controls={tabPanelId(tab.id)}
          aria-selected={tab.id === active}
          tabIndex={tab.id === active ? 0 : -1}
          className={tab.id === active ? 'ai-tabs__tab ai-tabs__tab--active' : 'ai-tabs__tab'}
          onClick={() => onChoose(tab.id)}
        >
          {t(tab.label)}
        </button>
      ))}
    </div>
  )
}

/** Both stay mounted: switching tabs keeps the answer and the chat where they were. */
function AiTabPanel({
  tab,
  active,
  children,
}: {
  readonly tab: AiTab
  readonly active: AiTab
  readonly children: React.ReactNode
}): React.JSX.Element {
  return (
    <div
      className="ai-pane__tab"
      role="tabpanel"
      id={tabPanelId(tab)}
      aria-labelledby={tabId(tab)}
      hidden={tab !== active}
    >
      {children}
    </div>
  )
}

interface AiActionsProps {
  readonly request: AiRequest
  readonly provider: AiProviderInfo | undefined
  readonly keepFormatting: boolean
  readonly readOnly: boolean
  readonly launch: AiLaunch | null
}

function AiActions({
  request,
  provider,
  keepFormatting,
  readOnly,
  launch,
}: AiActionsProps): React.JSX.Element {
  const t = useT()
  const [form, setForm] = useState<AiForm>(() => ({
    action: AiAction.Rewrite,
    targetLanguage: otherLanguage(),
    instruction: '',
  }))
  const [problem, setProblem] = useState<AiTargetProblem | null>(null)

  const generate = async (chosen: AiForm): Promise<void> => {
    if (provider === undefined) return
    setProblem(
      await request.start({
        providerId: provider.id,
        model: provider.defaultModel ?? '',
        action: chosen.action,
        targetLanguage: chosen.action === AiAction.Translate ? chosen.targetLanguage : undefined,
        instruction: chosen.action === AiAction.Custom ? chosen.instruction : undefined,
        format: responseFormat(chosen.action, keepFormatting),
      }),
    )
  }
  useLaunch(launch, provider, form, setForm, generate)

  return (
    <>
      <AiRequestForm
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
    </>
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

interface AiProviderFieldsProps {
  readonly providers: readonly AiProviderInfo[]
  readonly provider: AiProviderInfo | undefined
  readonly onProvider: (id: string) => void
  readonly onModel: (model: string) => void
  readonly running: boolean
}

function AiProviderFields(props: AiProviderFieldsProps): React.JSX.Element {
  const t = useT()
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
    </>
  )
}

interface AiRequestFormProps {
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
  const preview = useRef<HTMLDivElement>(null)
  const { state } = request
  if (state.phase === 'idle') return null
  const status = AI_STATUS_LABELS[state.phase]
  const author = authorOf(provider, state.model)
  const finished = state.phase === 'done' && state.text !== ''
  const markdown = state.format === AiFormat.Markdown
  const copy = (): void => {
    if (markdown) return copyContents(preview.current)
    output.current?.select()
    copySelection()
  }
  return (
    <>
      {markdown ? (
        <div className="popover__field">
          <span>{t('ai.panel.result')}</span>
          <AiMarkdownPreview ref={preview} markdown={state.text} label={t('ai.panel.result')} />
        </div>
      ) : (
        <label className="popover__field">
          <span>{t('ai.panel.result')}</span>
          <textarea ref={output} className="ai-pane__output" readOnly rows={10} value={state.text} />
        </label>
      )}
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
