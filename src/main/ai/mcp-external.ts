import { appliedOutcome, runDocumentTool } from '@services/ai/document-tools.js'
import { mcpAuthor } from '@services/ai/mcp-client-name.js'
import type { McpClient, McpToolResult } from '@services/ai/mcp-protocol.js'
import type { AiProposal } from '@services/ai/proposals.js'
import type { AiToolCall } from '@services/ai/provider.js'

/** The open document as the window has it now; each call asks again. */
export interface DocumentWindow {
  snapshot(): Promise<readonly string[]>
  /** False when the paragraph no longer reads as it did. */
  apply(proposal: AiProposal, author: string): Promise<boolean>
}

/** Its message is for the client: it says why the document did not answer. */
export class DocumentUnavailable extends Error {}

export type McpToolRunner = (call: AiToolCall, client: McpClient) => Promise<McpToolResult>

/** No limit on the number of proposals: each goes in as a revision or comment the person reviews. */
export function externalToolRunner(window: DocumentWindow, fallbackName: () => string): McpToolRunner {
  return async (call, client) => {
    try {
      const paragraphs = await window.snapshot()
      const outcome = runDocumentTool(call, { paragraphs, proposalCount: 0 })
      if (outcome.proposal === undefined) return outcome
      const applied = await window.apply(outcome.proposal, mcpAuthor(client.name, fallbackName()))
      return appliedOutcome(outcome.proposal, applied)
    } catch (cause) {
      if (cause instanceof DocumentUnavailable) return { content: cause.message, isError: true }
      throw cause
    }
  }
}
