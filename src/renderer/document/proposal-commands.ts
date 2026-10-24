import type { Editor } from '@tiptap/react'
import {
  AiProposalKind,
  type AiCommentProposal,
  type AiEditProposal,
  type AiInsertProposal,
  type AiProposal,
} from '@services/ai/proposals.js'
import { useWorkspace } from '../state/workspace.js'
import { buildTrackedEdits, type AiBlockInsertion } from './ai-apply.js'
import {
  locateInsertion,
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
const isInsert = (proposal: AiProposal): proposal is AiInsertProposal =>
  proposal.kind === AiProposalKind.Insert

interface LocatedInsertions {
  readonly located: readonly (AiBlockInsertion & { readonly id: string })[]
  readonly changed: readonly string[]
}

function locateInsertions(editor: Editor, inserts: readonly AiInsertProposal[]): LocatedInsertions {
  const located: (AiBlockInsertion & { readonly id: string })[] = []
  const changed: string[] = []
  for (const proposal of inserts) {
    const pos = locateInsertion(editor.state.doc, proposal)
    if (pos === null) changed.push(proposal.id)
    else located.push({ id: proposal.id, pos, markdown: proposal.content })
  }
  return { located, changed }
}

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

const idsOf = (located: readonly LocatedProposal<AiEditProposal | AiCommentProposal>[]): string[] =>
  located.map(({ proposal }) => proposal.id)

/**
 * Checked against the document as it is now. Edits, insertions and comments go in one transaction, one undo:
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
  const inserts = locateInsertions(editor, proposals.filter(isInsert))
  const textEdits = edits.located.map(({ proposal, range }) => ({ range, markdown: proposal.replace }))
  const tracked = buildTrackedEdits(editor.state, textEdits, inserts.located, author, new Date())
  const entries = comments.located.map(({ proposal, range }) => ({ range, text: proposal.comment }))
  const commented = insertCommentsAt(editor, entries, author, tracked ?? editor.state.tr)
  const applied = new Set([
    ...(tracked === null ? [] : [...idsOf(edits.located), ...inserts.located.map(({ id }) => id)]),
    ...(commented ? idsOf(comments.located) : []),
  ])
  return {
    applied: proposals.filter((proposal) => applied.has(proposal.id)).map((proposal) => proposal.id),
    changed: [...comments.changed, ...edits.changed, ...inserts.changed],
  }
}
