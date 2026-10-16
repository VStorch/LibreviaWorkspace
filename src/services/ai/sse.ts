import { MAX_AI_TOOL_ARGUMENTS_LENGTH, MAX_AI_TOOL_CALLS } from '@shared/limits.js'
import type { AiToolCall } from './provider.js'

/** A piece of a tool call: the id and name come in the first piece, the arguments in many. */
export interface OpenAiToolDelta {
  readonly index: number
  readonly id?: string
  readonly name?: string
  readonly arguments?: string
}

export type OpenAiEvent =
  | { readonly type: 'text'; readonly text: string }
  | { readonly type: 'tools'; readonly deltas: readonly OpenAiToolDelta[]; readonly text: string }
  | { readonly type: 'done' }
  | { readonly type: 'error'; readonly message: string }
  | { readonly type: 'refused' }
  | { readonly type: 'skip' }

const DATA_PREFIX = 'data:'
const DONE_MARKER = '[DONE]'
const CONTENT_FILTER = 'content_filter'
const SKIP: OpenAiEvent = { type: 'skip' }

interface ToolCallChunk {
  readonly index?: unknown
  readonly id?: unknown
  readonly function?: { readonly name?: unknown; readonly arguments?: unknown }
}

interface ChatDelta {
  readonly content?: unknown
  readonly tool_calls?: unknown
}

interface ChatChunk {
  readonly choices?: readonly {
    readonly delta?: ChatDelta
    readonly finish_reason?: unknown
  }[]
  readonly error?: { readonly message?: unknown }
}

/** One line of a `text/event-stream` from `/v1/chat/completions`. Comments and `event:` are skipped. */
export function parseOpenAiLine(line: string): OpenAiEvent {
  if (!line.startsWith(DATA_PREFIX)) return SKIP
  const data = line.slice(DATA_PREFIX.length).trim()
  if (data === DONE_MARKER) return { type: 'done' }

  let chunk: ChatChunk
  try {
    chunk = JSON.parse(data) as ChatChunk
  } catch {
    return SKIP
  }
  if (chunk.error !== undefined) {
    return { type: 'error', message: typeof chunk.error.message === 'string' ? chunk.error.message : '' }
  }
  const choice = chunk.choices?.[0]
  if (choice?.finish_reason === CONTENT_FILTER) return { type: 'refused' }
  return deltaEvent(choice?.delta ?? {})
}

function deltaEvent(delta: ChatDelta): OpenAiEvent {
  const text = typeof delta.content === 'string' ? delta.content : ''
  const deltas = toolDeltasOf(delta.tool_calls)
  if (deltas.length > 0) return { type: 'tools', deltas, text }
  return text !== '' ? { type: 'text', text } : SKIP
}

const stringOr = (value: unknown): string | undefined => (typeof value === 'string' ? value : undefined)

function toolDeltasOf(value: unknown): OpenAiToolDelta[] {
  if (!Array.isArray(value)) return []
  return (value as ToolCallChunk[]).map((chunk, position) => {
    const id = stringOr(chunk.id)
    const name = stringOr(chunk.function?.name)
    const args = stringOr(chunk.function?.arguments)
    return {
      index: typeof chunk.index === 'number' ? chunk.index : position,
      ...(id === undefined ? {} : { id }),
      ...(name === undefined ? {} : { name }),
      ...(args === undefined ? {} : { arguments: args }),
    }
  })
}

interface PartialCall {
  id: string
  name: string
  arguments: string
}

/** The calls of one reply, put together from their pieces in the order of their index. */
export class OpenAiToolCalls {
  private readonly calls = new Map<number, PartialCall>()

  /** `round` keeps the ids made up for calls without one apart from the earlier rounds'. */
  constructor(private readonly round = 0) {}

  /** False once the reply passes the calls or the argument length one reply may have. */
  add(deltas: readonly OpenAiToolDelta[]): boolean {
    for (const delta of deltas) {
      const call = this.calls.get(delta.index) ?? { id: '', name: '', arguments: '' }
      if (delta.id !== undefined && delta.id !== '') call.id = delta.id
      if (delta.name !== undefined) call.name += delta.name
      if (delta.arguments !== undefined) call.arguments += delta.arguments
      this.calls.set(delta.index, call)
      if (this.calls.size > MAX_AI_TOOL_CALLS || call.arguments.length > MAX_AI_TOOL_ARGUMENTS_LENGTH)
        return false
    }
    return true
  }

  /** Some servers leave the id out; the reply still needs one to pair each result with its call. */
  finished(): AiToolCall[] {
    return [...this.calls.entries()]
      .sort(([a], [b]) => a - b)
      .map(([index, call]) => ({
        ...call,
        id: call.id === '' ? `call_${this.round}_${index}` : call.id,
      }))
  }
}
