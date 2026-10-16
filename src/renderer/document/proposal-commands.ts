import type { Editor } from '@tiptap/react'
import {
  AiProposalKind,
  type AiCommentProposal,
  type AiEditProposal,
  type AiProposal,
} from '@services/ai/proposals.js'
import { useWorkspace } from '../state/workspace.js'
import { buildTrackedEdits } from './ai-apply.js'
import {
  locateProposal,
  locateProposals,
  type LocatedProposal,
  type LocatedProposals,
} from './ai-proposals.js'
import { insertCommentsAt } from './comment-commands.js'

export interface ProposalOutcome {
  readonly applied: readonly string[]
  /** The text they point at is no longer there as it was: nothing of them was applied. */
  readonly changed: readonly string[]
}

const isEdit = (proposal: AiProposal): proposal is AiEditProposal => proposal.kind === AiProposalKind.Edit
const isComment = (proposal: AiProposal): proposal is AiCommentProposal =>
  proposal.kind === AiProposalKind.Comment

function locateComments(
  editor: Editor,
  comments: readonly AiCommentProposal[],
): LocatedProposals<AiCommentProposal> {
  const located: LocatedProposal<AiCommentProposal>[] = []
  const changed: string[] = []
  for (const proposal of comments) {
    const range = locateProposal(editor.state.doc, proposal)
    if (range === null) changed.push(proposal.id)
    else located.push({ proposal, range })
  }
  return { located, changed }
}

const idsOf = (located: readonly LocatedProposal<AiProposal>[]): string[] =>
  located.map(({ proposal }) => proposal.id)

/**
 * Checked against the document as it is now. Edits and comments go in one transaction, one undo:
 * the comment ends are placed after the tracked edits, which keep the deleted text where it was.
 */
export function applyProposals(
  editor: Editor,
  proposals: readonly AiProposal[],
  author: string,
): ProposalOutcome {
  if (useWorkspace.getState().readOnly) return { applied: [], changed: [] }
  const edits = locateProposals(editor.state.doc, proposals.filter(isEdit))
  const comments = locateComments(editor, proposals.filter(isComment))
  const textEdits = edits.located.map(({ proposal, range }) => ({ range, text: proposal.replace }))
  const tracked = buildTrackedEdits(editor.state, textEdits, author, new Date())
  const entries = comments.located.map(({ proposal, range }) => ({ range, text: proposal.comment }))
  const commented = insertCommentsAt(editor, entries, author, tracked ?? editor.state.tr)
  const applied = new Set([
    ...(tracked === null ? [] : idsOf(edits.located)),
    ...(commented ? idsOf(comments.located) : []),
  ])
  return {
    applied: proposals.filter((proposal) => applied.has(proposal.id)).map((proposal) => proposal.id),
    changed: [...comments.changed, ...edits.changed],
  }
}
