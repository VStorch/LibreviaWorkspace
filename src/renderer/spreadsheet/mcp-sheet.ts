import { useEffect } from 'react'
import { create } from 'zustand'
import { runSheetTool, type SheetProposal } from '@services/ai/sheet-tools.js'
import { failure, type ToolOutcome } from '@services/ai/tools.js'
import type { Range } from '@services/spreadsheet/edit.js'
import type { McpDocumentRequest } from '@shared/api.js'
import { McpDocumentAsk, McpDocumentRefusal } from '@shared/types.js'
import { refused, type McpReply } from '../ai/mcp-reply.js'
import { useWorkspace } from '../state/workspace.js'

export interface PendingSheetProposal extends SheetProposal {
  readonly id: string
  readonly author: string
  /** The file it was made for: another file opened drops it. */
  readonly generation: number
}

interface PendingState {
  readonly pending: readonly PendingSheetProposal[]
}

/** Writes from MCP clients, oldest first, waiting for the person to apply or discard each. */
export const useSheetMcpProposals = create<PendingState>(() => ({ pending: [] }))

export function discardSheetProposal(id: string): void {
  useSheetMcpProposals.setState(({ pending }) => ({ pending: pending.filter((item) => item.id !== id) }))
}

function offer(proposal: SheetProposal, author: string, generation: number): void {
  const id = crypto.randomUUID()
  useSheetMcpProposals.setState(({ pending }) => ({
    pending: [
      ...pending.filter((item) => item.generation === generation),
      { ...proposal, id, author, generation },
    ],
  }))
}

/** The selection on the workbook on screen; null while a text document, a PDF or nothing is. */
let selection: Range | null = null

export function useOfferSheetToMcp(range: Range): void {
  useEffect(() => {
    selection = range
  }, [range])
  useEffect(
    () => () => {
      selection = null
    },
    [],
  )
}

const INCOMPLETE = 'The request came without the tool call or its author.'

const toolReply = (requestId: string, { content, isError }: ToolOutcome<unknown>): McpReply => ({
  requestId,
  ask: McpDocumentAsk.SheetTool,
  content,
  isError,
})

/** Reading works in read-only mode; a write only waits on screen, and never in read-only mode. */
export function answerSheetTool(request: McpDocumentRequest, selected = selection): McpReply {
  const { requestId, call, author } = request
  const { workbook, readOnly, generation } = useWorkspace.getState()
  if (workbook === null || selected === null) return refused(requestId, McpDocumentRefusal.NoSpreadsheet)
  if (call === undefined || author === undefined) return toolReply(requestId, failure(INCOMPLETE))
  const outcome = runSheetTool({ id: requestId, ...call }, { workbook, selection: selected })
  if (outcome.proposal !== undefined) {
    if (readOnly) return refused(requestId, McpDocumentRefusal.ReadOnly)
    offer(outcome.proposal, author, generation)
  }
  return toolReply(requestId, outcome)
}
