import { clipboard } from 'electron'
import { IpcChannel } from '@shared/ipc-channels.js'
import { mcpClientConfig, mcpHost, receiveMcpReply } from '../ai/mcp.js'
import { handle } from './registry.js'

/** The local MCP server's controls; whether it listens is the AI preferences' `mcpServer`. */
export function registerMcpHandlers(): void {
  handle(IpcChannel.McpDocumentReply, (payload, event) => ({
    received: receiveMcpReply(payload, event.sender),
  }))
  handle(IpcChannel.McpStatus, () => ({ clients: mcpHost().clients(), clientConfig: mcpClientConfig() }))
  handle(IpcChannel.McpDisconnect, () => {
    mcpHost().disconnectExternal()
    return { clients: [] }
  })
  handle(IpcChannel.McpCopyConfig, () => {
    clipboard.writeText(mcpClientConfig())
    return { copied: true as const }
  })
}
