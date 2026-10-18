export type ClaudeCliEvent =
  | { readonly type: 'text'; readonly text: string }
  | {
      readonly type: 'result'
      readonly failed: boolean
      readonly refused: boolean
      readonly costUsd: number | undefined
      readonly message: string
    }
  | { readonly type: 'model'; readonly model: string }
  /** Each round of a reply with tools is a message of its own. */
  | { readonly type: 'message'; readonly model: string | undefined }
  | { readonly type: 'skip' }

const SKIP: ClaudeCliEvent = { type: 'skip' }

interface CliLine {
  readonly type?: unknown
  readonly subtype?: unknown
  readonly model?: unknown
  readonly event?: {
    readonly type?: unknown
    readonly message?: { readonly model?: unknown }
    readonly delta?: { readonly type?: unknown; readonly text?: unknown }
  }
  readonly is_error?: unknown
  readonly stop_reason?: unknown
  readonly total_cost_usd?: unknown
  readonly result?: unknown
}

/** One line of `claude -p --output-format stream-json --include-partial-messages`. */
export function parseClaudeCliLine(line: string): ClaudeCliEvent {
  let parsed: CliLine
  try {
    parsed = JSON.parse(line) as CliLine
  } catch {
    return SKIP
  }
  if (parsed.type === 'system' && parsed.subtype === 'init') return modelOf(parsed.model)
  if (parsed.type === 'stream_event' && parsed.event?.type === 'message_start') {
    const model = parsed.event.message?.model
    return { type: 'message', model: typeof model === 'string' && model !== '' ? model : undefined }
  }
  if (parsed.type === 'stream_event') return textOf(parsed)
  if (parsed.type === 'result') return resultOf(parsed)
  return SKIP
}

function modelOf(model: unknown): ClaudeCliEvent {
  return typeof model === 'string' && model !== '' ? { type: 'model', model } : SKIP
}

function textOf(line: CliLine): ClaudeCliEvent {
  const event = line.event
  if (event?.type !== 'content_block_delta' || event.delta?.type !== 'text_delta') return SKIP
  const text = event.delta.text
  return typeof text === 'string' && text !== '' ? { type: 'text', text } : SKIP
}

function resultOf(line: CliLine): ClaudeCliEvent {
  return {
    type: 'result',
    failed: line.is_error === true,
    refused: line.stop_reason === 'refusal',
    costUsd: typeof line.total_cost_usd === 'number' ? line.total_cost_usd : undefined,
    message: typeof line.result === 'string' ? line.result : '',
  }
}
