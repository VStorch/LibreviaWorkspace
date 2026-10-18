import { useEffect } from 'react'
import { create } from 'zustand'
import { runSheetTool, type SheetProposal } from '@services/ai/sheet-tools.js'
import { failure, type ToolOutcome } from '@services/ai/tools.js'
import type { Range } from '@services/spreadsheet/edit.js'
import type { McpDocumentRequest } from '@shared/api.js'
import type { IpcRequest } from '@shared/ipc.js'
import type { IpcChannel } from '@shared/ipc-channels.js'
import { MCP_DOCUMENT_REFUSED, McpDocumentAsk, McpDocumentRefusal } from '@shared/types.js'
import { useWorkspace } from '../state/workspace.js'

type Reply = IpcRequest<typeof IpcChannel.McpDocumentReply>

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

/** Until the workbook leaves the screen. */
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

const NO_CALL = 'The request came without a tool call.'

const toolReply = (requestId: string, { content, isError }: ToolOutcome<unknown>): Reply => ({
  requestId,
  ask: McpDocumentAsk.SheetTool,
  content,
  isError,
})

const refused = (requestId: string, refusal: McpDocumentRefusal): Reply => ({
  requestId,
  ask: MCP_DOCUMENT_REFUSED,
  refusal,
})

/** Reading works in read-only mode; a write only waits on screen, and never in read-only mode. */
export function answerSheetTool(request: McpDocumentRequest): Reply {
  const { requestId, call, author } = request
  const { workbook, readOnly, generation } = useWorkspace.getState()
  if (workbook === null || selection === null) return refused(requestId, McpDocumentRefusal.NoSpreadsheet)
  if (call === undefined || author === undefined) return toolReply(requestId, failure(NO_CALL))
  const outcome = runSheetTool({ id: requestId, ...call }, { workbook, selection })
  if (outcome.proposal !== undefined) {
    if (readOnly) return refused(requestId, McpDocumentRefusal.ReadOnly)
    offer(outcome.proposal, author, generation)
  }
  return toolReply(requestId, outcome)
}
