import { randomBytes } from 'node:crypto'
import { chmod, mkdir, readFile, rm, writeFile } from 'node:fs/promises'
import { join } from 'node:path'

export const MCP_BRIDGE_FLAG = '--mcp-bridge'
/** An internal session's key: it takes the place of the file's. */
export const MCP_TOKEN_ENV = 'LIBREVIA_MCP_TOKEN'

const MCP_FOLDER = 'mcp'
const SOCKET_FILE = 'mcp.sock'
const TOKEN_FILE = 'token'
const PIPE_NAME_FILE = 'pipe'
const PIPE_PREFIX = '\\\\.\\pipe\\librevia-mcp-'
const KEY_BYTES = 32
const PIPE_NAME_BYTES = 16
const PRIVATE_FOLDER_MODE = 0o700
const PRIVATE_FILE_MODE = 0o600

export interface McpPaths {
  readonly folder: string
  readonly token: string
  /** Unix: the socket itself. Windows: the file that names the pipe, random on each opening. */
  readonly address: string
  readonly windows: boolean
}

export function mcpPaths(userData: string, platform: NodeJS.Platform = process.platform): McpPaths {
  const folder = join(userData, MCP_FOLDER)
  const windows = platform === 'win32'
  return {
    folder,
    token: join(folder, TOKEN_FILE),
    address: join(folder, windows ? PIPE_NAME_FILE : SOCKET_FILE),
    windows,
  }
}

export const isMcpBridge = (argv: readonly string[]): boolean => argv.includes(MCP_BRIDGE_FLAG)

export const newMcpKey = (): string => randomBytes(KEY_BYTES).toString('hex')

export const newPipeName = (): string => PIPE_PREFIX + randomBytes(PIPE_NAME_BYTES).toString('hex')

/** `mkdir` keeps the mode of a folder that already exists: `chmod` corrects it. */
export async function ensurePrivateFolder(folder: string): Promise<void> {
  await mkdir(folder, { recursive: true, mode: PRIVATE_FOLDER_MODE })
  await chmod(folder, PRIVATE_FOLDER_MODE)
}

export async function writePrivateFile(path: string, content: string): Promise<void> {
  await rm(path, { force: true })
  await writeFile(path, content, { mode: PRIVATE_FILE_MODE, flag: 'wx' })
}

export const restrictToOwner = (path: string): Promise<void> => chmod(path, PRIVATE_FILE_MODE)

/** Where the bridge connects: on Windows, the pipe the app opened last. */
export async function listeningAddress(paths: McpPaths): Promise<string> {
  return paths.windows ? (await readFile(paths.address, 'utf8')).trim() : paths.address
}
