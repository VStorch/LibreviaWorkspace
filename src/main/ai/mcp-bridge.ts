import { readFile } from 'node:fs/promises'
import { connect, Socket } from 'node:net'
import type { Readable, Writable } from 'node:stream'
import { app } from 'electron'
import { listeningAddress, mcpPaths, MCP_TOKEN_ENV, type McpPaths } from './mcp-paths.js'

const BRIDGE_FAILED_EXIT_CODE = 1
const STDIN_FD = 0
const NOT_LISTENING =
  'Librevia is not open, or its preference that lets external programs read the document is off.'

export interface BridgeEnds {
  readonly paths: McpPaths
  /** An internal session's key, from the environment: it takes the place of the file's. */
  readonly key: string | undefined
  readonly input: Readable
  readonly output: Writable
}

/** A dumb channel: the key first, then stdin to the socket and the socket to stdout, until one side ends. */
export async function bridge({ paths, key, input, output }: BridgeEnds): Promise<void> {
  const secret = key ?? (await readFile(paths.token, 'utf8')).trim()
  const socket = connect(await listeningAddress(paths))
  await new Promise<void>((resolve, reject) => {
    socket.once('connect', resolve)
    socket.once('error', reject)
  })
  socket.write(`${secret}\n`)
  input.pipe(socket)
  socket.pipe(output)
  await new Promise<void>((resolve) => {
    socket.once('close', () => resolve())
    input.once('end', () => socket.end())
  })
}

/** No window, no renderer: the same data folder as the app, so the same socket and key. */
export function runMcpBridge(): void {
  app.dock?.hide()
  const key = process.env[MCP_TOKEN_ENV]
  const ends = {
    paths: mcpPaths(app.getPath('userData')),
    key: key === undefined || key === '' ? undefined : key,
    // In Electron's main process `process.stdin` is an empty stream that ends at once, which would
    // close the channel before the client says anything. The descriptor itself works, and a socket
    // on it, unlike a file stream, leaves no blocked read to hold the exit back.
    input: new Socket({ fd: STDIN_FD, readable: true, writable: false }),
    output: process.stdout,
  }
  bridge(ends).then(
    () => app.exit(0),
    () => {
      process.stderr.write(`${NOT_LISTENING}\n`)
      app.exit(BRIDGE_FAILED_EXIT_CODE)
    },
  )
}
