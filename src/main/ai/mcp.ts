import { app, type WebContents } from 'electron'
import { MCP_SERVER_NAME } from '@services/ai/mcp-protocol.js'
import { IpcChannel } from '@shared/ipc-channels.js'
import { t } from '../i18n.js'
import { broadcastPush } from '../window.js'
import { McpHost, type McpBridgeLaunch } from './mcp-host.js'
import { MCP_BRIDGE_FLAG, mcpPaths } from './mcp-paths.js'
import { RendererDocument } from './mcp-window.js'

const CONFIG_INDENT = 2
const USER_DATA_SWITCH = '--user-data-dir'

/** What a client runs: the AppImage itself when packaged as one, since its mount path changes. */
export function mcpBridgeLaunch(): McpBridgeLaunch {
  if (app.isPackaged) {
    return { command: process.env['APPIMAGE'] ?? process.execPath, args: [MCP_BRIDGE_FLAG] }
  }
  return { command: process.execPath, args: [app.getAppPath(), MCP_BRIDGE_FLAG] }
}

/** To paste into the client's configuration, as Claude Desktop and Claude Code read it. */
export function mcpClientConfig(): string {
  const { command, args } = mcpBridgeLaunch()
  return JSON.stringify({ mcpServers: { [MCP_SERVER_NAME]: { command, args } } }, null, CONFIG_INDENT)
}

const documentWindow = new RendererDocument()
let host: McpHost | null = null

/** The internal sessions' bridge reaches this very data folder, wherever the app was started with. */
function internalBridge(): McpBridgeLaunch {
  const { command, args } = mcpBridgeLaunch()
  return { command, args: [...args, `${USER_DATA_SWITCH}=${app.getPath('userData')}`] }
}

export function mcpHost(): McpHost {
  host ??= new McpHost({
    paths: mcpPaths(app.getPath('userData')),
    version: app.getVersion(),
    bridge: internalBridge(),
    document: documentWindow,
    fallbackName: () => t('ai.mcp.externalClient'),
    onClientsChanged: (clients) => broadcastPush(IpcChannel.McpClients, { clients: [...clients] }),
  })
  return host
}

export function attachMcpWindow(contents: WebContents): void {
  documentWindow.attach(contents)
}

export function receiveMcpReply(...args: Parameters<RendererDocument['receive']>): boolean {
  return documentWindow.receive(...args)
}

export function disposeMcp(): void {
  void host?.dispose().catch(() => undefined)
}
