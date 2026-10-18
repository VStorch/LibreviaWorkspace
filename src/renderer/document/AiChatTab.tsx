import { useCallback, useEffect, useRef, useState, type RefObject } from 'react'
import { AiRole } from '@services/ai/provider.js'
import { MAX_AI_CHAT_MESSAGE_LENGTH } from '@shared/limits.js'
import { useT } from '../i18n.js'
import { copyContents } from './ai-copy.js'
import { AiChatProposals } from './AiChatProposals.js'
import { AiMarkdownPreview } from './AiMarkdownPreview.js'
import type { AiChat, AiChatMessage } from './useAiChat.js'
import { AI_STATUS_LABELS } from '../ai/ai-status.js'
import { authorOf, type AiProviderInfo } from './useAiProviders.js'

interface AiChatTabBaseProps {
  readonly chat: AiChat
  readonly provider: AiProviderInfo
  readonly readOnly: boolean
}

interface AiChatLogProps extends AiChatTabBaseProps {
  readonly logRef: RefObject<HTMLDivElement | null>
  /** After Apply or Discard, whose button leaves with the proposal's state. */
  readonly onProposalSettled: () => void
}

interface AiChatTabProps extends AiChatTabBaseProps {
  readonly proposals: boolean
  readonly onProposals: (on: boolean) => void
}

function AiProposalsChoice({
  proposals,
  onProposals,
}: Pick<AiChatTabProps, 'proposals' | 'onProposals'>): React.JSX.Element {
  const t = useT()
  return (
    <label className="popover__check">
      <input type="checkbox" checked={proposals} onChange={(event) => onProposals(event.target.checked)} />
      {t('ai.chat.proposals')}
    </label>
  )
}

/** The buttons the panel keeps the focus on; `data-proposal-action` marks them in AiChatProposals. */
const PROPOSAL_ACTION = '[data-proposal-action]:not(:disabled)'

/**
 * The focus stays in the panel: on the next proposal's buttons, or on the message field once none
 * is left. The editor only gets it when something is inserted there.
 */
function useFocusAfterProposal(field: RefObject<HTMLTextAreaElement | null>): {
  readonly logRef: RefObject<HTMLDivElement | null>
  readonly settled: () => void
} {
  const logRef = useRef<HTMLDivElement>(null)
  const [settledCount, setSettledCount] = useState(0)
  useEffect(() => {
    if (settledCount === 0) return
    const next = logRef.current?.querySelector<HTMLButtonElement>(PROPOSAL_ACTION)
    ;(next ?? field.current)?.focus()
  }, [settledCount, field])
  const settled = useCallback(() => setSettledCount((count) => count + 1), [])
  return { logRef, settled }
}

export function AiChatTab({
  chat,
  provider,
  readOnly,
  proposals,
  onProposals,
}: AiChatTabProps): React.JSX.Element {
  const t = useT()
  const [draft, setDraft] = useState('')
  const field = useRef<HTMLTextAreaElement>(null)
  const { logRef, settled } = useFocusAfterProposal(field)
  const canSend = !chat.running && draft.trim() !== ''
  const submit = (): void => {
    if (!canSend) return
    setDraft('')
    field.current?.focus()
    void chat.send(draft, {
      providerId: provider.id,
      model: provider.defaultModel ?? '',
      proposals,
    })
  }
  const restart = (): void => {
    chat.clear()
    field.current?.focus()
  }
  return (
    <>
      <AiChatLog
        chat={chat}
        provider={provider}
        readOnly={readOnly}
        logRef={logRef}
        onProposalSettled={settled}
      />
      <label className="popover__field">
        <span>{t('ai.chat.message')}</span>
        <textarea
          ref={field}
          rows={3}
          value={draft}
          maxLength={MAX_AI_CHAT_MESSAGE_LENGTH}
          placeholder={t('ai.chat.placeholder')}
          onChange={(event) => setDraft(event.target.value)}
          onKeyDown={(event) => {
            if (event.key !== 'Enter' || event.shiftKey || event.nativeEvent.isComposing) return
            event.preventDefault()
            submit()
          }}
        />
      </label>
      {draft.length >= MAX_AI_CHAT_MESSAGE_LENGTH && (
        <p className="ai-pane__warning" role="status">
          {t('ai.chat.tooLong', { max: MAX_AI_CHAT_MESSAGE_LENGTH })}
        </p>
      )}
      <div className="ai-pane__actions">
        {chat.running ? (
          <button type="button" className="btn" onClick={chat.cancel}>
            {t('ai.panel.stop')}
          </button>
        ) : (
          <button type="button" className="btn btn--primary" disabled={!canSend} onClick={submit}>
            {t('ai.chat.send')}
          </button>
        )}
        <button type="button" className="btn" disabled={chat.messages.length === 0} onClick={restart}>
          {t('ai.chat.new')}
        </button>
      </div>
      <AiProposalsChoice proposals={proposals} onProposals={onProposals} />
      <p className="popover__hint">{t('ai.chat.hint')}</p>
    </>
  )
}

function AiChatLog({
  chat,
  provider,
  readOnly,
  logRef,
  onProposalSettled,
}: AiChatLogProps): React.JSX.Element {
  const t = useT()
  return (
    <div ref={logRef} className="ai-chat" role="log" aria-live="polite" aria-label={t('ai.chat.log')}>
      {chat.messages.length === 0 && <p className="ai-pane__note">{t('ai.chat.empty')}</p>}
      {chat.messages.map((message) =>
        message.role === AiRole.User ? (
          <div key={message.id} className="ai-chat__message ai-chat__message--person">
            <span className="ai-chat__speaker">{t('ai.chat.you')}</span>
            <p className="ai-chat__text">{message.text}</p>
          </div>
        ) : (
          <AiChatReply
            key={message.id}
            message={message}
            readOnly={readOnly}
            onInsert={() => chat.insertBelow(message, authorOf(provider, message.model))}
            onApply={(ids) => {
              chat.apply(message.id, ids, authorOf(provider, message.model))
              onProposalSettled()
            }}
            onDiscard={(id) => {
              chat.discard(message.id, id)
              onProposalSettled()
            }}
          />
        ),
      )}
    </div>
  )
}

function AiChatReply({
  message,
  readOnly,
  onInsert,
  onApply,
  onDiscard,
}: {
  readonly message: AiChatMessage
  readonly readOnly: boolean
  readonly onInsert: () => void
  readonly onApply: (ids: readonly string[]) => void
  readonly onDiscard: (id: string) => void
}): React.JSX.Element {
  const t = useT()
  const preview = useRef<HTMLDivElement>(null)
  const status = AI_STATUS_LABELS[message.phase]
  const finished = message.phase === 'done' && message.text !== ''
  return (
    <div className="ai-chat__message" aria-busy={message.phase === 'running'}>
      <span className="ai-chat__speaker">{t('ai.chat.assistant')}</span>
      {message.text !== '' && (
        <AiMarkdownPreview ref={preview} markdown={message.text} label={t('ai.panel.result')} />
      )}
      {status !== undefined && (
        <p className="ai-pane__note" role="status">
          {t(status)}
        </p>
      )}
      {message.error !== undefined && <p className="ai-pane__warning">{message.error}</p>}
      {message.costUsd !== undefined && (
        <p className="ai-pane__note">{t('ai.panel.cost', { cost: message.costUsd.toFixed(4) })}</p>
      )}
      {message.proposals !== undefined && message.proposals.length > 0 && (
        <AiChatProposals
          proposals={message.proposals}
          readOnly={readOnly}
          onApply={onApply}
          onDiscard={onDiscard}
        />
      )}
      {finished && (
        <div className="ai-pane__actions">
          <button type="button" className="btn" disabled={readOnly} onClick={onInsert}>
            {t('ai.panel.insertBelow')}
          </button>
          <button type="button" className="btn" onClick={() => copyContents(preview.current)}>
            {t('ai.panel.copy')}
          </button>
        </div>
      )}
    </div>
  )
}
