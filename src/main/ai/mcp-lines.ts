import type { Socket } from 'node:net'
import { takeLines } from '@services/ai/lines.js'
import { JsonRpcErrorCode, mcpErrorLine, type McpConnection } from '@services/ai/mcp-protocol.js'
import { MAX_AI_STREAM_LINE_LENGTH } from '@shared/limits.js'

const TOO_LONG = mcpErrorLine(JsonRpcErrorCode.InvalidRequest, 'Message too long: it was discarded.')

/** One message per line, answered in order. A line over the limit is answered with an error and dropped. */
export function serveLines(socket: Socket, connection: McpConnection, initial: string): void {
  let buffer = ''
  let discarding = false
  let queue: Promise<void> = Promise.resolve()

  const send = (reply: string | null): void => {
    if (reply !== null && !socket.destroyed) socket.write(`${reply}\n`)
  }
  const answer = (line: string): void => {
    queue = queue.then(() => connection.handle(line)).then(send, () => undefined)
  }
  const take = (chunk: string): void => {
    buffer += chunk
    if (discarding) {
      const end = buffer.indexOf('\n')
      buffer = end === -1 ? '' : buffer.slice(end + 1)
      discarding = end === -1
    }
    const split = takeLines(buffer)
    for (const line of split.lines) {
      if (line.length > MAX_AI_STREAM_LINE_LENGTH) send(TOO_LONG)
      else answer(line)
    }
    buffer = split.rest
    if (buffer.length > MAX_AI_STREAM_LINE_LENGTH) {
      send(TOO_LONG)
      buffer = ''
      discarding = true
    }
  }

  socket.on('data', take)
  if (initial !== '') take(initial)
}
