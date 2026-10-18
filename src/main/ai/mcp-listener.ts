import { createServer, type Server, type Socket } from 'node:net'
import { rm } from 'node:fs/promises'
import {
  ensurePrivateFolder,
  newPipeName,
  restrictToOwner,
  writePrivateFile,
  type McpPaths,
} from './mcp-paths.js'

function listenOn(server: Server, path: string): Promise<void> {
  return new Promise((resolve, reject) => {
    server.once('error', reject)
    server.listen(path, () => {
      server.off('error', reject)
      resolve()
    })
  })
}

/**
 * Only a path, never a port: a Unix socket in a folder only the owner enters, or a named pipe whose
 * random name only the owner can read. Starting and stopping go one after the other, since closing
 * a socket removes its file, which could be the new one's.
 */
export class McpListener {
  private server: Server | null = null
  private transition: Promise<void> = Promise.resolve()

  constructor(
    private readonly paths: McpPaths,
    private readonly accept: (socket: Socket) => void,
  ) {}

  start(): Promise<void> {
    return this.serialize(() => this.listen())
  }

  stop(): Promise<void> {
    return this.serialize(() => this.close())
  }

  private serialize(step: () => Promise<void>): Promise<void> {
    const next = this.transition.then(step)
    this.transition = next.catch(() => undefined)
    return next
  }

  private async listen(): Promise<void> {
    if (this.server !== null) return
    await ensurePrivateFolder(this.paths.folder)
    const address = this.paths.windows ? newPipeName() : this.paths.address
    if (!this.paths.windows) await rm(address, { force: true })
    const server = createServer(this.accept)
    await listenOn(server, address)
    if (this.paths.windows) await writePrivateFile(this.paths.address, address)
    else await restrictToOwner(address)
    this.server = server
  }

  /** The connections are ended by whoever owns them; the socket stops taking new ones at once. */
  private close(): Promise<void> {
    this.server?.close()
    this.server = null
    return this.paths.windows ? rm(this.paths.address, { force: true }) : Promise.resolve()
  }
}
