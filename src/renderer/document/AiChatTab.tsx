import { useRef, useState } from 'react'
import { AiRole } from '@services/ai/provider.js'
import type { MessageKey } from '@shared/i18n/index.js'
import { MAX_AI_CHAT_MESSAGE_LENGTH } from '@shared/limits.js'
import { useT } from '../i18n.js'
import { copyContents } from './ai-copy.js'
import { AiMarkdownPreview } from './AiMarkdownPreview.js'
import type { AiChat, AiChatMessage } from './useAiChat.js'
import type { AiPhase } from './useAiRequest.js'
import { authorOf, type AiProviderInfo } from './useAiProviders.js'

export const AI_STATUS_LABELS: Partial<Record<AiPhase, MessageKey>> = {
  running: 'ai.panel.running',
  canceled: 'ai.panel.canceled',
  declined: 'ai.panel.declined',
  failed: 'ai.panel.failed',
}

interface AiChatTabProps {
  readonly chat: AiChat
  readonly provider: AiProviderInfo
  readonly readOnly: boolean
}

export function AiChatTab({ chat, provider, readOnly }: AiChatTabProps): React.JSX.Element {
  const t = useT()
  const [draft, setDraft] = useState('')
  const field = useRef<HTMLTextAreaElement>(null)
  const canSend = !chat.running && draft.trim() !== ''
  const submit = (): void => {
    if (!canSend) return
    setDraft('')
    field.current?.focus()
    void chat.send(draft, { providerId: provider.id, model: provider.defaultModel ?? '' })
  }
  const restart = (): void => {
    chat.clear()
    field.current?.focus()
  }
  return (
    <>
      <AiChatLog chat={chat} provider={provider} readOnly={readOnly} />
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
      <p className="popover__hint">{t('ai.chat.hint')}</p>
    </>
  )
}

function AiChatLog({ chat, provider, readOnly }: AiChatTabProps): React.JSX.Element {
  const t = useT()
  return (
    <div className="ai-chat" role="log" aria-live="polite" aria-label={t('ai.chat.log')}>
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
}: {
  readonly message: AiChatMessage
  readonly readOnly: boolean
  readonly onInsert: () => void
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
      {message.costUsd !== undefined && (
        <p className="ai-pane__note">{t('ai.panel.cost', { cost: message.costUsd.toFixed(4) })}</p>
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
