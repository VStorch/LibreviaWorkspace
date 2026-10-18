import { timingSafeEqual } from 'node:crypto'
import { rm } from 'node:fs/promises'
import type { Socket } from 'node:net'
import { DOCUMENT_TOOLS } from '@services/ai/document-tools.js'
import { EXTERNAL_INSTRUCTIONS, EXTERNAL_TOOLS } from '@services/ai/external-tools.js'
import { mcpClientName } from '@services/ai/mcp-client-name.js'
import { createMcpConnection, MCP_SERVER_NAME, type McpServerOptions } from '@services/ai/mcp-protocol.js'
import { MAX_MCP_CLIENTS } from '@shared/limits.js'
import { externalToolRunner, type DocumentWindow } from './mcp-external.js'
import { serveLines } from './mcp-lines.js'
import { McpListener } from './mcp-listener.js'
import {
  ensurePrivateFolder,
  MCP_TOKEN_ENV,
  newMcpKey,
  writePrivateFile,
  type McpPaths,
} from './mcp-paths.js'
import { runToolCall, type ToolDocument, type ToolRun } from './tool-loop.js'
import type { ToolSession, ToolSessions } from './tool-server.js'

export const MCP_AUTH_TIMEOUT_MS = 2_000
/** The key is 64 hexadecimal digits: anything much longer before the newline is not one. */
const MAX_KEY_LINE_LENGTH = 128

export interface McpBridgeLaunch {
  readonly command: string
  readonly args: readonly string[]
}

export interface McpHostOptions {
  readonly paths: McpPaths
  readonly version: string
  /** How a client starts the bridge; an internal session's comes with this app's data folder. */
  readonly bridge: McpBridgeLaunch
  readonly document: DocumentWindow
  /** "External assistant", in the app's language. */
  readonly fallbackName: () => string
  readonly onClientsChanged: (clients: readonly string[]) => void
  readonly authTimeoutMs?: number
}

interface InternalSession {
  readonly run: ToolRun
  readonly sockets: Set<Socket>
}

type Role = { readonly internal: InternalSession } | { readonly external: true }

/** Equal lengths first: `timingSafeEqual` throws on different ones. */
function keysMatch(given: string, expected: string): boolean {
  const a = Buffer.from(given)
  const b = Buffer.from(expected)
  return a.length === b.length && timingSafeEqual(a, b)
}

/**
 * The local MCP server. It listens only while the preference is on or a conversation of Claude
 * Code is running, and serves only those who present a key: the file's, while the preference is
 * on, or one internal session's, which reads only that message's snapshot.
 */
export class McpHost implements ToolSessions {
  private readonly listener: McpListener
  private readonly internal = new Map<string, InternalSession>()
  private readonly external = new Map<Socket, string>()
  private externalKey: string | null = null

  constructor(private readonly options: McpHostOptions) {
    this.listener = new McpListener(options.paths, (socket) => this.accept(socket))
  }

  /** A new key each time it is turned on: a client that had the old one must read the file again. */
  async setExternal(enabled: boolean): Promise<void> {
    if (enabled === (this.externalKey !== null)) return
    const { paths } = this.options
    if (enabled) {
      const key = newMcpKey()
      this.externalKey = key
      await ensurePrivateFolder(paths.folder)
      await writePrivateFile(paths.token, key)
      await this.listener.start()
      return
    }
    this.externalKey = null
    this.disconnectExternal()
    await rm(paths.token, { force: true })
    await this.stopIfIdle()
  }

  disconnectExternal(): void {
    for (const socket of this.external.keys()) socket.destroy()
  }

  clients(): string[] {
    return [...this.external.values()].map((name) => mcpClientName(name, this.options.fallbackName()))
  }

  async open(document: ToolDocument): Promise<ToolSession> {
    const key = newMcpKey()
    const session: InternalSession = { run: { document, proposals: [], read: 0 }, sockets: new Set() }
    this.internal.set(key, session)
    try {
      await this.listener.start()
    } catch (cause) {
      this.closeSession(key)
      throw cause
    }
    const { command, args } = this.options.bridge
    return {
      server: {
        name: MCP_SERVER_NAME,
        command,
        args,
        env: { [MCP_TOKEN_ENV]: key },
        tools: DOCUMENT_TOOLS.map((tool) => tool.name),
      },
      proposals: () => session.run.proposals,
      close: () => this.closeSession(key),
    }
  }

  async dispose(): Promise<void> {
    this.externalKey = null
    this.disconnectExternal()
    for (const key of [...this.internal.keys()]) this.closeSession(key)
    await this.listener.stop()
  }

  private closeSession(key: string): void {
    const session = this.internal.get(key)
    if (session === undefined) return
    this.internal.delete(key)
    for (const socket of session.sockets) socket.destroy()
    void this.stopIfIdle().catch(() => undefined)
  }

  private stopIfIdle(): Promise<void> {
    const idle = this.externalKey === null && this.internal.size === 0
    return idle ? this.listener.stop() : Promise.resolve()
  }

  private accept(socket: Socket): void {
    socket.on('error', () => undefined)
    socket.setEncoding('utf8')
    const timer = setTimeout(() => socket.destroy(), this.options.authTimeoutMs ?? MCP_AUTH_TIMEOUT_MS)
    socket.once('close', () => clearTimeout(timer))
    let buffer = ''
    const onData = (chunk: string): void => {
      buffer += chunk
      const end = buffer.indexOf('\n')
      if (end === -1) {
        if (buffer.length > MAX_KEY_LINE_LENGTH) socket.destroy()
        return
      }
      clearTimeout(timer)
      socket.off('data', onData)
      this.admit(socket, buffer.slice(0, end).trim(), buffer.slice(end + 1))
    }
    socket.on('data', onData)
  }

  private admit(socket: Socket, key: string, rest: string): void {
    const role = this.roleOf(key)
    if (role === null) socket.destroy()
    else if ('internal' in role) this.serveInternal(socket, role.internal, rest)
    else this.serveExternal(socket, rest)
  }

  /** Every key is compared, so the time taken says nothing about which one matched. */
  private roleOf(key: string): Role | null {
    let role: Role | null = null
    for (const [candidate, session] of this.internal) {
      if (keysMatch(key, candidate)) role = { internal: session }
    }
    const external = this.externalKey !== null && keysMatch(key, this.externalKey)
    if (role !== null || !external) return role
    return this.external.size < MAX_MCP_CLIENTS ? { external: true } : null
  }

  private serveInternal(socket: Socket, session: InternalSession, rest: string): void {
    session.sockets.add(socket)
    socket.once('close', () => session.sockets.delete(socket))
    this.serve(socket, rest, {
      tools: DOCUMENT_TOOLS,
      runTool: (call) => Promise.resolve(runToolCall(call, session.run)),
    })
  }

  private serveExternal(socket: Socket, rest: string): void {
    this.external.set(socket, '')
    socket.once('close', () => {
      this.external.delete(socket)
      this.notify()
    })
    this.notify()
    this.serve(socket, rest, {
      tools: EXTERNAL_TOOLS,
      instructions: EXTERNAL_INSTRUCTIONS,
      runTool: externalToolRunner(this.options.document, this.options.fallbackName),
      onInitialize: (client) => {
        if (!this.external.has(socket)) return
        this.external.set(socket, client.name)
        this.notify()
      },
    })
  }

  private serve(socket: Socket, rest: string, server: Omit<McpServerOptions, 'version'>): void {
    serveLines(socket, createMcpConnection({ ...server, version: this.options.version }), rest)
  }

  private notify(): void {
    this.options.onClientsChanged(this.clients())
  }
}
