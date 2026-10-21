import { describe, expect, it } from 'vitest'
import { mcpAuthor, mcpClientName } from './mcp-client-name.js'

describe('MCP client name', () => {
  it('the author always carries the (MCP) suffix', () => {
    expect(mcpAuthor('Claude Code', 'Assistente externo')).toBe('Claude Code (MCP)')
  })

  it('strips control characters and surrounding spaces', () => {
    expect(mcpAuthor('  Cli\u0000en\u001bte\n ', 'Assistente externo')).toBe('Cliente (MCP)')
  })

  it('strips whole terminal color sequences', () => {
    expect(mcpClientName('Evil\u0007\u001b[31mBot\u001b[0m', 'x')).toBe('EvilBot')
  })

  it('truncates at 64 characters', () => {
    expect(mcpClientName('a'.repeat(100), 'x')).toBe('a'.repeat(64))
  })

  it('an empty name becomes the catalog phrase, still with the suffix', () => {
    expect(mcpAuthor(' \u0007 ', 'Assistente externo')).toBe('Assistente externo (MCP)')
  })
})
