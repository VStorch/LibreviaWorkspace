import { mkdtemp, readFile, rm, stat } from 'node:fs/promises'
import { connect, Server, type Socket } from 'node:net'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import type { AiProposal } from '@services/ai/proposals.js'
import type { DocumentWindow } from './mcp-external.js'
import { McpHost, type McpHostOptions } from './mcp-host.js'
import { mcpPaths, MCP_TOKEN_ENV } from './mcp-paths.js'

const IS_WINDOWS = process.platform === 'win32'
const SHORT_AUTH_MS = 100

interface Client {
  readonly socket: Socket
  readonly closed: Promise<void>
  call(method: string, params?: unknown): Promise<Record<string, unknown>>
}

function client(address: string, key: string): Promise<Client> {
  const socket = connect(address)
  socket.setEncoding('utf8')
  let buffer = ''
  let nextId = 1
  const waiting = new Map<number, (message: Record<string, unknown>) => void>()
  socket.on('data', (chunk: string) => {
    buffer += chunk
    const lines = buffer.split('\n')
    buffer = lines.pop() ?? ''
    for (const line of lines) {
      const message = JSON.parse(line) as Record<string, unknown>
      waiting.get(message['id'] as number)?.(message)
    }
  })
  socket.on('error', () => undefined)
  const closed = new Promise<void>((resolve) => socket.once('close', () => resolve()))
  return new Promise((resolve) => {
    socket.once('connect', () => {
      if (key !== '') socket.write(`${key}\n`)
      resolve({
        socket,
        closed,
        call: (method, params) => {
          const id = nextId++
          socket.write(`${JSON.stringify({ jsonrpc: '2.0', id, method, params })}\n`)
          return new Promise((answer) => waiting.set(id, answer))
        },
      })
    })
  })
}

const textOf = (message: Record<string, unknown>): string =>
  (message['result'] as { content: { text: string }[] }).content[0]!.text

describe('local MCP server', () => {
  let userData: string
  let host: McpHost
  let applied: { proposal: AiProposal; author: string }[]
  let sheetCalls: { name: string; author: string }[]
  let clientsSeen: (readonly string[])[]
  const paths = (): ReturnType<typeof mcpPaths> => mcpPaths(userData)
  const address = async (): Promise<string> =>
    IS_WINDOWS ? (await readFile(paths().address, 'utf8')).trim() : paths().address
  const externalKey = (): Promise<string> => readFile(paths().token, 'utf8')

  const document: DocumentWindow = {
    snapshot: () => Promise.resolve(['Primeiro parágrafo.', 'O gato dorme.']),
    apply: (proposal, author) => {
      applied.push({ proposal, author })
      return Promise.resolve(true)
    },
    runSheetTool: (call, author) => {
      sheetCalls.push({ name: call.name, author })
      return Promise.resolve({ content: 'Sheet: Vendas', isError: false })
    },
  }

  function makeHost(overrides: Partial<McpHostOptions> = {}): McpHost {
    return new McpHost({
      paths: paths(),
      version: '0.0.0',
      bridge: { command: '/bin/librevia', args: ['--mcp-bridge'] },
      document,
      fallbackName: () => 'Assistente externo',
      onClientsChanged: (clients) => clientsSeen.push(clients),
      ...overrides,
    })
  }

  beforeEach(async () => {
    userData = await mkdtemp(join(tmpdir(), 'librevia-mcp-'))
    applied = []
    sheetCalls = []
    clientsSeen = []
    host = makeHost()
  })

  afterEach(async () => {
    vi.restoreAllMocks()
    await host.dispose()
    await rm(userData, { recursive: true, force: true })
  })

  it('listens only on a path, never on a port', async () => {
    const listen = vi.spyOn(Server.prototype, 'listen')
    await host.setExternal(true)
    expect(listen).toHaveBeenCalledTimes(1)
    expect(typeof listen.mock.calls[0]![0]).toBe('string')
    expect(listen.mock.calls[0]![0]).toBe(await address())
  })

  it.skipIf(IS_WINDOWS)('folder 0700, socket and key 0600', async () => {
    await host.setExternal(true)
    const mode = async (path: string): Promise<number> => (await stat(path)).mode & 0o777
    expect(await mode(paths().folder)).toBe(0o700)
    expect(await mode(paths().address)).toBe(0o600)
    expect(await mode(paths().token)).toBe(0o600)
  })

  it('wrong key closes the connection', async () => {
    await host.setExternal(true)
    const wrong = await client(await address(), 'f'.repeat(64))
    await wrong.closed
    expect(wrong.socket.destroyed).toBe(true)
  })

  it('without the key within the timeout, closes', async () => {
    await host.dispose()
    host = makeHost({ authTimeoutMs: SHORT_AUTH_MS })
    await host.setExternal(true)
    const silent = await client(await address(), '')
    await silent.closed
    expect(silent.socket.destroyed).toBe(true)
  })

  it('the key changes each time the preference is turned on; when off, refuses the external client', async () => {
    await host.setExternal(true)
    const first = await externalKey()
    const connected = await client(await address(), first)
    await connected.call('ping')
    await host.setExternal(false)
    await connected.closed
    await expect(stat(paths().token)).rejects.toThrow()
    await host.setExternal(true)
    expect(await externalKey()).not.toBe(first)
    const stale = await client(await address(), first)
    await stale.closed
  })

  it('the third external connection is refused', async () => {
    await host.setExternal(true)
    const key = await externalKey()
    const one = await client(await address(), key)
    const two = await client(await address(), key)
    await one.call('ping')
    await two.call('ping')
    const three = await client(await address(), key)
    await three.closed
    expect(clientsSeen.at(-1)).toHaveLength(2)
  })

  it('the external client shows by name, and the proposal goes in at once with the (MCP) author', async () => {
    await host.setExternal(true)
    const external = await client(await address(), await externalKey())
    await external.call('initialize', { clientInfo: { name: 'Editor X' } })
    expect(host.clients()).toEqual(['Editor X'])
    const answer = await external.call('tools/call', {
      name: 'propose_edit',
      arguments: { paragraph: 2, find: 'gato', replace: 'cão' },
    })
    expect(textOf(answer)).toContain('tracked revision')
    expect(applied).toHaveLength(1)
    expect(applied[0]!.author).toBe('Editor X (MCP)')
    host.disconnectExternal()
    await external.closed
    expect(clientsSeen.at(-1)).toEqual([])
  })

  it('the external client has the document and sheet tools, which go to the window', async () => {
    await host.setExternal(true)
    const external = await client(await address(), await externalKey())
    const start = await external.call('initialize', { clientInfo: { name: 'Editor X' } })
    expect(start['result']).toMatchObject({ instructions: expect.stringContaining('spreadsheet') })
    const list = await external.call('tools/list')
    const names = (list['result'] as { tools: { name: string }[] }).tools.map((tool) => tool.name)
    expect(names).toEqual([
      'read_paragraphs',
      'search_document',
      'propose_edit',
      'add_comment',
      'list_sheets',
      'read_cells',
      'find_cells',
      'write_cells',
    ])
    const read = await external.call('tools/call', { name: 'read_cells', arguments: { range: 'A1' } })
    expect(textOf(read)).toBe('Sheet: Vendas')
    expect(sheetCalls).toEqual([{ name: 'read_cells', author: 'Editor X (MCP)' }])
    const unknown = await external.call('tools/call', { name: 'run_shell', arguments: {} })
    expect(textOf(unknown)).toContain(
      'read_paragraphs, search_document, propose_edit, add_comment, list_sheets',
    )
  })

  it('the internal session reads only its own snapshot and dies when done', async () => {
    const mine = await host.open({ paragraphs: ['Só meu.'], readBudget: 1_000 })
    const other = await host.open({ paragraphs: ['De outro.'], readBudget: 1_000 })
    const key = mine.server.env[MCP_TOKEN_ENV]!
    const session = await client(await address(), key)
    const read = await session.call('tools/call', {
      name: 'read_paragraphs',
      arguments: { from: 1, count: 5 },
    })
    expect(textOf(read)).toBe('[1] Só meu.')
    await session.call('tools/call', {
      name: 'add_comment',
      arguments: { paragraph: 1, find: 'meu', comment: 'Bom.' },
    })
    expect(mine.proposals()).toHaveLength(1)
    expect(applied).toEqual([])
    const sheet = await session.call('tools/call', { name: 'read_cells', arguments: { range: 'A1' } })
    expect(textOf(sheet)).toContain('There is no tool named "read_cells"')
    expect(sheetCalls).toEqual([])
    mine.close()
    await session.closed
    const again = await client(await address(), key)
    await again.closed
    other.close()
  })

  it('without preference or session, does not listen', async () => {
    const session = await host.open({ paragraphs: [], readBudget: 1 })
    session.close()
    await host.setExternal(false)
    await vi.waitFor(async () => {
      if (!IS_WINDOWS) await expect(stat(paths().address)).rejects.toThrow()
    })
  })
})
