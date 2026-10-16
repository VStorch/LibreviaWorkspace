import { useId } from 'react'
import { AiProposalKind } from '@services/ai/proposals.js'
import type { MessageKey } from '@shared/i18n/index.js'
import { useT } from '../i18n.js'
import type { ChatProposal, ProposalState } from './useAiChat.js'

const STATE_LABELS: Partial<Record<ProposalState, MessageKey>> = {
  applied: 'ai.proposals.applied',
  discarded: 'ai.proposals.discarded',
  changed: 'ai.proposals.changed',
}

interface AiChatProposalsProps {
  readonly proposals: readonly ChatProposal[]
  readonly readOnly: boolean
  readonly onApply: (ids: readonly string[]) => void
  readonly onDiscard: (id: string) => void
}

export function AiChatProposals({
  proposals,
  readOnly,
  onApply,
  onDiscard,
}: AiChatProposalsProps): React.JSX.Element {
  const t = useT()
  const open = proposals.filter((proposal) => proposal.state === 'open')
  return (
    <div className="ai-proposals" role="group" aria-label={t('ai.proposals.title')}>
      {proposals.map((proposal) => (
        <AiProposalCard
          key={proposal.id}
          proposal={proposal}
          readOnly={readOnly}
          onApply={() => onApply([proposal.id])}
          onDiscard={() => onDiscard(proposal.id)}
        />
      ))}
      {open.length > 1 && (
        <div className="ai-pane__actions">
          <button
            type="button"
            className="btn"
            data-proposal-action=""
            disabled={readOnly}
            onClick={() => onApply(open.map((proposal) => proposal.id))}
          >
            {t('ai.proposals.applyAll')}
          </button>
        </div>
      )}
    </div>
  )
}

function AiProposalBody({ proposal }: { readonly proposal: ChatProposal }): React.JSX.Element {
  const t = useT()
  if (proposal.kind === AiProposalKind.Comment) {
    return (
      <>
        <q className="ai-proposal__find">{proposal.find}</q>
        <p className="ai-proposal__comment">{proposal.comment}</p>
      </>
    )
  }
  return (
    <>
      <del className="ai-proposal__find">{proposal.find}</del>
      {proposal.replace === '' ? (
        <span className="ai-pane__note">{t('ai.proposals.removal')}</span>
      ) : (
        <ins className="ai-proposal__replace">{proposal.replace}</ins>
      )}
      {proposal.reason !== undefined && <p className="ai-pane__note">{proposal.reason}</p>}
    </>
  )
}

function AiProposalCard({
  proposal,
  readOnly,
  onApply,
  onDiscard,
}: {
  readonly proposal: ChatProposal
  readonly readOnly: boolean
  readonly onApply: () => void
  readonly onDiscard: () => void
}): React.JSX.Element {
  const t = useT()
  const status = STATE_LABELS[proposal.state]
  const where = proposal.kind === AiProposalKind.Edit ? 'ai.proposals.edit' : 'ai.proposals.comment'
  const describedBy = useId()
  return (
    <div className="ai-proposal">
      <div id={describedBy} className="ai-proposal__text">
        <span className="ai-chat__speaker">{t(where, { paragraph: proposal.paragraph })}</span>
        <AiProposalBody proposal={proposal} />
      </div>
      {status !== undefined ? (
        <p className="ai-pane__note" role="status">
          {t(status)}
        </p>
      ) : (
        <div className="ai-pane__actions">
          <button
            type="button"
            className="btn"
            data-proposal-action=""
            aria-describedby={describedBy}
            disabled={readOnly}
            onClick={onApply}
          >
            {t('ai.proposals.apply')}
          </button>
          <button
            type="button"
            className="btn"
            data-proposal-action=""
            aria-describedby={describedBy}
            onClick={onDiscard}
          >
            {t('ai.proposals.discard')}
          </button>
        </div>
      )}
    </div>
  )
}
