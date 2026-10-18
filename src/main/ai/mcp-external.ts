import { appliedOutcome, isDocumentTool, runDocumentTool } from '@services/ai/document-tools.js'
import { mcpAuthor } from '@services/ai/mcp-client-name.js'
import type { McpClient, McpToolResult } from '@services/ai/mcp-protocol.js'
import type { AiProposal } from '@services/ai/proposals.js'
import type { AiToolCall } from '@services/ai/provider.js'
import { isSheetTool } from '@services/ai/sheet-tools.js'
import { EXTERNAL_TOOLS } from '@services/ai/external-tools.js'

/** The open document as the window has it now; each call asks again. */
export interface DocumentWindow {
  snapshot(): Promise<readonly string[]>
  /** False when the paragraph no longer reads as it did. */
  apply(proposal: AiProposal, author: string): Promise<boolean>
  /** Runs over the workbook on screen; a write waits there for the person. */
  runSheetTool(call: AiToolCall, author: string): Promise<McpToolResult>
}

/** Its message is for the client: it says why the document did not answer. */
export class DocumentUnavailable extends Error {}

export type McpToolRunner = (call: AiToolCall, client: McpClient) => Promise<McpToolResult>

const unknownTool = (name: string): McpToolResult => ({
  content: `There is no tool named "${name}". The tools are: ${EXTERNAL_TOOLS.map((tool) => tool.name).join(', ')}.`,
  isError: true,
})

async function runOnDocument(
  window: DocumentWindow,
  call: AiToolCall,
  author: string,
): Promise<McpToolResult> {
  const paragraphs = await window.snapshot()
  const outcome = runDocumentTool(call, { paragraphs, proposalCount: 0 })
  if (outcome.proposal === undefined) return outcome
  return appliedOutcome(outcome.proposal, await window.apply(outcome.proposal, author))
}

/**
 * No limit on the number of proposals: in a text document each goes in as a revision or comment
 * the person reviews; in a spreadsheet each waits for the person to apply it.
 */
export function externalToolRunner(window: DocumentWindow, fallbackName: () => string): McpToolRunner {
  return async (call, client) => {
    const author = mcpAuthor(client.name, fallbackName())
    try {
      if (isDocumentTool(call.name)) return await runOnDocument(window, call, author)
      if (isSheetTool(call.name)) return await window.runSheetTool(call, author)
      return unknownTool(call.name)
    } catch (cause) {
      if (cause instanceof DocumentUnavailable) return { content: cause.message, isError: true }
      throw cause
    }
  }
}
