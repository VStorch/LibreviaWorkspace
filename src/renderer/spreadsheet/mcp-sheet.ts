import { useEffect } from 'react'
import { create } from 'zustand'
import { runSheetTool, type SheetProposal } from '@services/ai/sheet-tools.js'
import { failure, type ToolOutcome } from '@services/ai/tools.js'
import type { Range } from '@services/spreadsheet/edit.js'
import type { McpDocumentRequest } from '@shared/api.js'
import { McpDocumentAsk, McpDocumentRefusal } from '@shared/types.js'
import { refused, type McpReply } from '../ai/mcp-reply.js'
import { activeWorkspaceStore, createTabSlot, useOwnTabId, useTabs } from '../state/tabs.js'

export interface PendingSheetProposal extends SheetProposal {
  readonly id: string
  readonly author: string
  /** The tab and file it was made for: another file opened in that tab drops it. */
  readonly tabId: string
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

export function isProposalFor(item: PendingSheetProposal, tabId: string, generation: number): boolean {
  return item.tabId === tabId && item.generation === generation
}

function offer(proposal: SheetProposal, author: string, generation: number): void {
  const id = crypto.randomUUID()
  const { activeId: tabId } = useTabs.getState()
  useSheetMcpProposals.setState(({ pending }) => ({
    pending: [
      ...pending.filter((item) => item.tabId !== tabId || item.generation === generation),
      { ...proposal, id, author, tabId, generation },
    ],
  }))
}

/** The selection on the active tab's workbook; none while it shows a text document, a PDF or nothing. */
const selection = createTabSlot<Range>()

export function useOfferSheetToMcp(range: Range): void {
  const tabId = useOwnTabId()
  useEffect(() => selection.hold(tabId, range), [tabId, range])
}

const INCOMPLETE = 'The request came without the tool call or its author.'

const toolReply = (requestId: string, { content, isError }: ToolOutcome<unknown>): McpReply => ({
  requestId,
  ask: McpDocumentAsk.SheetTool,
  content,
  isError,
})

/** Reading works in read-only mode; a write only waits on screen, and never in read-only mode. */
export function answerSheetTool(request: McpDocumentRequest, selected = selection.active()): McpReply {
  const { requestId, call, author } = request
  const { workbook, readOnly, generation } = activeWorkspaceStore().getState()
  if (workbook === null || selected === undefined) return refused(requestId, McpDocumentRefusal.NoSpreadsheet)
  if (call === undefined || author === undefined) return toolReply(requestId, failure(INCOMPLETE))
  const outcome = runSheetTool({ id: requestId, ...call }, { workbook, selection: selected })
  if (outcome.proposal !== undefined) {
    if (readOnly) return refused(requestId, McpDocumentRefusal.ReadOnly)
    offer(outcome.proposal, author, generation)
  }
  return toolReply(requestId, outcome)
}
