import type { AiProposal } from '@services/ai/proposals.js'
import type { AiPrompt, AiProvider, AiToolServer } from '@services/ai/provider.js'
import { relay } from './relay.js'
import type { ToolDocument, ToolLoopResult } from './tool-loop.js'

/** One message's access to the document tools: its own key, its own snapshot. */
export interface ToolSession {
  readonly server: AiToolServer
  proposals(): readonly AiProposal[]
  /** Revokes the key and ends the connections it opened. */
  close(): void
}

export interface ToolSessions {
  open(document: ToolDocument): Promise<ToolSession>
}

/**
 * For a provider that runs the tool loop itself, through the local MCP server. The proposals come
 * back with the reply, as with the loop run here; the session ends with it, canceled or not.
 */
export function relayThroughToolServer(sessions: ToolSessions, document: ToolDocument) {
  return async (
    provider: AiProvider,
    prompt: AiPrompt,
    signal: AbortSignal,
    add: (text: string) => void,
  ): Promise<ToolLoopResult> => {
    const session = await sessions.open(document)
    try {
      const relayed = await relay(provider, { ...prompt, toolServer: session.server }, signal, add)
      return {
        ...(relayed.costUsd === undefined ? {} : { costUsd: relayed.costUsd }),
        ...(relayed.model === undefined ? {} : { model: relayed.model }),
        proposals: [...session.proposals()],
      }
    } finally {
      session.close()
    }
  }
}
