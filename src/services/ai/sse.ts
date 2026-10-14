export type OpenAiEvent =
  | { readonly type: 'text'; readonly text: string }
  | { readonly type: 'done' }
  | { readonly type: 'error'; readonly message: string }
  | { readonly type: 'refused' }
  | { readonly type: 'skip' }

const DATA_PREFIX = 'data:'
const DONE_MARKER = '[DONE]'
const CONTENT_FILTER = 'content_filter'
const SKIP: OpenAiEvent = { type: 'skip' }

interface ChatChunk {
  readonly choices?: readonly {
    readonly delta?: { readonly content?: unknown }
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
  const content = choice?.delta?.content
  return typeof content === 'string' && content !== '' ? { type: 'text', text: content } : SKIP
}
