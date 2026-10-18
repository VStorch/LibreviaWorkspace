import { describe, expect, it } from 'vitest'
import { mcpAuthor, mcpClientName } from './mcp-client-name.js'

describe('nome do cliente MCP', () => {
  it('o autor leva sempre o sufixo (MCP)', () => {
    expect(mcpAuthor('Claude Code', 'Assistente externo')).toBe('Claude Code (MCP)')
  })

  it('tira caracteres de controle e espaços das pontas', () => {
    expect(mcpAuthor('  Cli\u0000en\u001bte\n ', 'Assistente externo')).toBe('Cliente (MCP)')
  })

  it('tira as sequências de cor do terminal inteiras', () => {
    expect(mcpClientName('Evil\u0007\u001b[31mBot\u001b[0m', 'x')).toBe('EvilBot')
  })

  it('corta em 64 caracteres', () => {
    expect(mcpClientName('a'.repeat(100), 'x')).toBe('a'.repeat(64))
  })

  it('nome vazio vira a frase do catálogo, ainda com o sufixo', () => {
    expect(mcpAuthor(' \u0007 ', 'Assistente externo')).toBe('Assistente externo (MCP)')
  })
})
