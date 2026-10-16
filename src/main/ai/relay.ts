import type { AiPrompt, AiProvider, AiToolCall } from '@services/ai/provider.js'
import { MAX_AI_RESPONSE_LENGTH } from '@shared/limits.js'

export interface RelayMeta {
  readonly costUsd?: number
  readonly model?: string
}

export interface Relayed extends RelayMeta {
  readonly text: string
  readonly toolCalls: readonly AiToolCall[]
}

/** Shared by the rounds of one reply: the person gets at most one response's worth of text. */
export interface RelayBudget {
  used: number
}

export const budgetSpent = (budget: RelayBudget): boolean => budget.used >= MAX_AI_RESPONSE_LENGTH

export async function relay(
  provider: AiProvider,
  prompt: AiPrompt,
  signal: AbortSignal,
  add: (text: string) => void,
  budget: RelayBudget = { used: 0 },
): Promise<Relayed> {
  let meta: RelayMeta = {}
  let text = ''
  const toolCalls: AiToolCall[] = []
  for await (const part of provider.stream(prompt, signal)) {
    if ('toolCall' in part) {
      toolCalls.push(part.toolCall)
      continue
    }
    if (!('text' in part)) {
      meta = { ...meta, ...part }
      continue
    }
    const piece = part.text.slice(0, MAX_AI_RESPONSE_LENGTH - budget.used)
    budget.used += piece.length
    text += piece
    if (piece !== '') add(piece)
    if (budgetSpent(budget)) break
  }
  signal.throwIfAborted()
  return { ...meta, text, toolCalls }
}
