import { describe, expect, it, vi } from 'vitest'
import { DOCUMENT_TOOLS } from './document-tools.js'
import { createMcpConnection, JsonRpcErrorCode, type McpServerOptions } from './mcp-protocol.js'

function connection(overrides: Partial<McpServerOptions> = {}) {
  return createMcpConnection({
    version: '1.2.3',
    tools: DOCUMENT_TOOLS,
    runTool: () => Promise.resolve({ content: 'feito', isError: false }),
    ...overrides,
  })
}

const request = (id: number, method: string, params?: unknown): string =>
  JSON.stringify({ jsonrpc: '2.0', id, method, ...(params === undefined ? {} : { params }) })

async function reply(handle: Promise<string | null>): Promise<Record<string, unknown>> {
  const text = await handle
  expect(text).not.toBeNull()
  return JSON.parse(text!) as Record<string, unknown>
}

describe('protocolo MCP', () => {
  it('initialize devolve a versão pedida quando conhecida, as capacidades e o nome do servidor', async () => {
    const onInitialize = vi.fn()
    const answer = await reply(
      connection({ onInitialize }).handle(
        request(1, 'initialize', { protocolVersion: '2025-03-26', clientInfo: { name: 'Cliente' } }),
      ),
    )
    expect(answer).toEqual({
      jsonrpc: '2.0',
      id: 1,
      result: {
        protocolVersion: '2025-03-26',
        capabilities: { tools: {} },
        serverInfo: { name: 'librevia', version: '1.2.3' },
      },
    })
    expect(onInitialize).toHaveBeenCalledWith({ name: 'Cliente' })
  })

  it('versão desconhecida recebe a mais nova', async () => {
    const answer = await reply(
      connection().handle(request(1, 'initialize', { protocolVersion: '1999-01-01' })),
    )
    expect(answer['result']).toMatchObject({ protocolVersion: '2025-06-18' })
  })

  it('tools/list lista as ferramentas com o esquema de entrada', async () => {
    const answer = await reply(connection().handle(request(2, 'tools/list')))
    const tools = (answer['result'] as { tools: { name: string; inputSchema: unknown }[] }).tools
    expect(tools.map((tool) => tool.name)).toEqual(DOCUMENT_TOOLS.map((tool) => tool.name))
    expect(tools[0]!.inputSchema).toEqual(DOCUMENT_TOOLS[0]!.parameters)
  })

  it('tools/call passa os argumentos como JSON e o cliente que se apresentou', async () => {
    const runTool = vi.fn(() => Promise.resolve({ content: 'não há', isError: true }))
    const subject = connection({ runTool })
    await subject.handle(request(1, 'initialize', { clientInfo: { name: 'Editor X' } }))
    const answer = await reply(
      subject.handle(request(3, 'tools/call', { name: 'read_paragraphs', arguments: { from: 1, count: 2 } })),
    )
    expect(runTool).toHaveBeenCalledWith(
      { id: '3', name: 'read_paragraphs', arguments: '{"from":1,"count":2}' },
      { name: 'Editor X' },
    )
    expect(answer['result']).toEqual({ content: [{ type: 'text', text: 'não há' }], isError: true })
  })

  it('ferramenta que falha vira resultado de erro, sem derrubar a conexão', async () => {
    const answer = await reply(
      connection({ runTool: () => Promise.reject(new Error('x')) }).handle(
        request(4, 'tools/call', { name: 'read_paragraphs' }),
      ),
    )
    expect(answer['result']).toMatchObject({ isError: true })
  })

  it('responde ping e ignora notificações', async () => {
    const subject = connection()
    expect(await reply(subject.handle(request(5, 'ping')))).toEqual({ jsonrpc: '2.0', id: 5, result: {} })
    expect(
      await subject.handle(JSON.stringify({ jsonrpc: '2.0', method: 'notifications/initialized' })),
    ).toBeNull()
  })

  it('método desconhecido, JSON inválido, parâmetros inválidos e lote recebem o erro certo', async () => {
    const subject = connection()
    const codeOf = async (line: string): Promise<unknown> =>
      ((await reply(subject.handle(line)))['error'] as { code: number }).code
    expect(await codeOf(request(6, 'resources/list'))).toBe(JsonRpcErrorCode.MethodNotFound)
    expect(await codeOf('{ não é json')).toBe(JsonRpcErrorCode.ParseError)
    expect(await codeOf(request(7, 'tools/call', { name: 3 }))).toBe(JsonRpcErrorCode.InvalidParams)
    expect(await codeOf(JSON.stringify([JSON.parse(request(8, 'ping'))]))).toBe(
      JsonRpcErrorCode.InvalidRequest,
    )
  })

  it('a resposta do cliente a um pedido não é respondida', async () => {
    expect(await connection().handle(JSON.stringify({ jsonrpc: '2.0', id: 9, result: {} }))).toBeNull()
  })
})
