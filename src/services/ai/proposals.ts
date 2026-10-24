/** What the renderer needs of the chat's proposals, apart from the tools that make them. */
export const AiProposalKind = { Edit: 'edit', Comment: 'comment', Insert: 'insert' } as const

export type AiProposalKind = (typeof AiProposalKind)[keyof typeof AiProposalKind]

interface ProposalPlace {
  readonly id: string
  /** One-based, as numbered in the prompt. */
  readonly paragraph: number
  /** The paragraph's whole text when proposed: applying needs it unchanged, wherever it moved. */
  readonly paragraphText: string
}

interface ProposalAnchor extends ProposalPlace {
  /** Appears exactly once in the paragraph's text without deletions. */
  readonly find: string
}

export interface AiEditProposal extends ProposalAnchor {
  readonly kind: typeof AiProposalKind.Edit
  /** Markdown. */
  readonly replace: string
  readonly reason?: string | undefined
}

export interface AiCommentProposal extends ProposalAnchor {
  readonly kind: typeof AiProposalKind.Comment
  readonly comment: string
}

/** New blocks after the paragraph; paragraph zero is the start of the document, with no text. */
export interface AiInsertProposal extends ProposalPlace {
  readonly kind: typeof AiProposalKind.Insert
  /** Markdown. */
  readonly content: string
}

/** What the model asked for; nothing in the document changes until the person applies it. */
export type AiProposal = AiEditProposal | AiCommentProposal | AiInsertProposal

export type AiAnchoredProposal = AiEditProposal | AiCommentProposal

export function occurrences(text: string, find: string): number {
  let count = 0
  for (let at = text.indexOf(find); at !== -1; at = text.indexOf(find, at + find.length)) count++
  return count
}
