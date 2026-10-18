import { MAX_MCP_CLIENT_NAME_LENGTH } from '@shared/limits.js'

const TERMINAL_SEQUENCES = /\p{Cc}\[[0-?]*[ -/]*[@-~]/gu
const CONTROL_CHARACTERS = /\p{Cc}/gu

/** Marks every change from outside the app: the person always knows who wrote it. */
export const MCP_AUTHOR_SUFFIX = ' (MCP)'

/** The name the client gave itself on `initialize`, which it chooses freely. */
export function mcpClientName(name: string, fallback: string): string {
  const clean = Array.from(name.replace(TERMINAL_SEQUENCES, '').replace(CONTROL_CHARACTERS, '').trim())
    .slice(0, MAX_MCP_CLIENT_NAME_LENGTH)
    .join('')
    .trim()
  return clean === '' ? fallback : clean
}

export function mcpAuthor(name: string, fallback: string): string {
  return mcpClientName(name, fallback) + MCP_AUTHOR_SUFFIX
}
