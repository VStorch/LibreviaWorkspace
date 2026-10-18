import { useEffect } from 'react'
import type { Editor } from '@tiptap/react'
import { closeHistory } from '@tiptap/pm/history'
import type { IpcRequest } from '@shared/ipc.js'
import { IpcChannel } from '@shared/ipc-channels.js'
import { pushContracts } from '@shared/ipc.js'
import type { McpDocumentRequest } from '@shared/api.js'
import { MCP_DOCUMENT_REFUSED, McpDocumentAsk, McpDocumentRefusal } from '@shared/types.js'
import { useWorkspace } from '../state/workspace.js'
import { chatParagraphs } from './ai-proposals.js'
import { applyProposals } from './proposal-commands.js'

type Reply = IpcRequest<typeof IpcChannel.McpDocumentReply>

/** The text document on screen; none while a spreadsheet, a PDF or nothing is open. */
let shown: Editor | null = null

/** Until the editor leaves the screen. */
export function useOfferToMcp(editor: Editor | null): void {
  useEffect(() => {
    if (editor === null) return undefined
    shown = editor
    return () => {
      if (shown === editor) shown = null
    }
  }, [editor])
}

const refused = (requestId: string, refusal: McpDocumentRefusal): Reply => ({
  requestId,
  ask: MCP_DOCUMENT_REFUSED,
  refusal,
})

/** Always a tracked revision or a comment, each its own undo step, whatever the editing mode. */
function applyOne(editor: Editor, request: McpDocumentRequest): boolean {
  if (request.proposal === undefined || request.author === undefined) return false
  editor.view.dispatch(closeHistory(editor.state.tr))
  const outcome = applyProposals(editor, [request.proposal], request.author)
  editor.view.dispatch(closeHistory(editor.state.tr))
  return outcome.applied.length > 0
}

function answer(request: McpDocumentRequest): Reply {
  const { requestId } = request
  if (Date.now() > request.deadline) return refused(requestId, McpDocumentRefusal.Expired)
  const editor = shown
  if (editor === null || editor.isDestroyed) return refused(requestId, McpDocumentRefusal.NoTextDocument)
  if (request.ask === McpDocumentAsk.Snapshot) {
    return { requestId, ask: McpDocumentAsk.Snapshot, paragraphs: chatParagraphs(editor.state.doc) }
  }
  if (useWorkspace.getState().readOnly) return refused(requestId, McpDocumentRefusal.ReadOnly)
  return { requestId, ask: McpDocumentAsk.Apply, applied: applyOne(editor, request) }
}

/** Returns the unsubscribe function. */
export function answerMcpRequests(): () => void {
  return window.api.mcp.onDocumentRequest((payload) => {
    // The second zod check, which the sandboxed preload cannot do.
    const parsed = pushContracts[IpcChannel.McpDocumentRequest].safeParse(payload)
    if (parsed.success) void window.api.mcp.reply(answer(parsed.data))
  })
}
