import { DOCUMENT_TOOLS, runDocumentTool } from '@services/ai/document-tools.js'
import type { AiProposal } from '@services/ai/proposals.js'
import type { AiPrompt, AiProvider, AiToolCall, AiToolExchange, AiToolResult } from '@services/ai/provider.js'
import { MAX_AI_TOOL_CALLS, MAX_AI_TOOL_ROUNDS } from '@shared/limits.js'
import { budgetSpent, relay, type RelayBudget, type RelayMeta } from './relay.js'

export interface ToolLoopResult extends RelayMeta {
  readonly proposals: readonly AiProposal[]
}

const ROUND_SEPARATOR = '\n\n'

/** The text of a new round reads as a new paragraph after the previous round's. */
function separated(add: (text: string) => void, budget: RelayBudget): (text: string) => void {
  let separator = budget.used > 0 ? ROUND_SEPARATOR : ''
  return (text) => {
    add(separator + text)
    separator = ''
  }
}

/** The snapshot the tools read, and how much of it one reply may take back to the model. */
export interface ToolDocument {
  readonly paragraphs: readonly string[]
  readonly readBudget: number
}

interface ToolRun {
  readonly document: ToolDocument
  readonly proposals: AiProposal[]
  /** Characters of results already sent back, all rounds together. */
  read: number
}

const tooManyCalls = (callId: string): AiToolResult => ({
  callId,
  content: `Only ${MAX_AI_TOOL_CALLS} tool calls are answered per turn; call fewer at a time.`,
  isError: true,
})

const READ_LIMIT_REACHED =
  'The reading limit for this message is used up: answer with what you have already read.'

function runCall(call: AiToolCall, run: ToolRun): AiToolResult {
  const { paragraphs, readBudget } = run.document
  const outcome = runDocumentTool(call, { paragraphs, proposalCount: run.proposals.length })
  if (!outcome.isError && run.read + outcome.content.length > readBudget) {
    return { callId: call.id, content: READ_LIMIT_REACHED, isError: true }
  }
  run.read += outcome.content.length
  if (outcome.proposal !== undefined) run.proposals.push(outcome.proposal)
  return { callId: call.id, content: outcome.content, isError: outcome.isError }
}

function runCalls(exchange: Omit<AiToolExchange, 'results'>, run: ToolRun): AiToolExchange {
  const results = exchange.calls.map((call, index) =>
    index < MAX_AI_TOOL_CALLS ? runCall(call, run) : tooManyCalls(call.id),
  )
  return { ...exchange, results }
}

function sumCost(a: number | undefined, b: number | undefined): number | undefined {
  if (a === undefined) return b
  return b === undefined ? a : a + b
}

/**
 * Each round streams the model's text and runs the tools it called over the snapshot, until a
 * round calls none. Past the last round the calls go unanswered and the reply ends as it is.
 */
export async function relayWithTools(
  provider: AiProvider,
  prompt: AiPrompt,
  document: ToolDocument,
  signal: AbortSignal,
  add: (text: string) => void,
): Promise<ToolLoopResult> {
  const budget: RelayBudget = { used: 0 }
  const exchanges: AiToolExchange[] = []
  const run: ToolRun = { document, proposals: [], read: 0 }
  let meta: RelayMeta = {}
  for (let round = 0; round < MAX_AI_TOOL_ROUNDS; round++) {
    const request = { ...prompt, tools: DOCUMENT_TOOLS, exchanges: [...exchanges] }
    const { text, toolCalls, ...roundMeta } = await relay(
      provider,
      request,
      signal,
      separated(add, budget),
      budget,
    )
    const costUsd = sumCost(meta.costUsd, roundMeta.costUsd)
    meta = { ...meta, ...roundMeta, ...(costUsd === undefined ? {} : { costUsd }) }
    if (toolCalls.length === 0 || budgetSpent(budget)) break
    exchanges.push(runCalls({ text, calls: toolCalls }, run))
  }
  return { ...meta, proposals: run.proposals }
}
