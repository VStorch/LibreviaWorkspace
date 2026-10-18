import { create } from 'zustand'
import { IpcChannel } from '@shared/ipc-channels.js'
import { pushContracts } from '@shared/ipc.js'

/** The programs connected to the local MCP server now, by the names they gave. */
interface McpClientsState {
  readonly clients: readonly string[]
}

export const useMcpClients = create<McpClientsState>(() => ({ clients: [] }))

/** Returns the unsubscribe function. */
export function watchMcpClients(): () => void {
  const stop = window.api.mcp.onClients((payload) => {
    // The second zod check, which the sandboxed preload cannot do.
    const parsed = pushContracts[IpcChannel.McpClients].safeParse(payload)
    if (parsed.success) useMcpClients.setState({ clients: parsed.data.clients })
  })
  void window.api.mcp.status({}).then((result) => {
    if (result.ok) useMcpClients.setState({ clients: result.data.clients })
  })
  return stop
}

export async function disconnectMcpClients(): Promise<void> {
  const result = await window.api.mcp.disconnect({})
  if (result.ok) useMcpClients.setState({ clients: result.data.clients })
}
