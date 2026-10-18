import type { IpcRequest } from '@shared/ipc.js'
import type { IpcChannel } from '@shared/ipc-channels.js'
import { MCP_DOCUMENT_REFUSED, type McpDocumentRefusal } from '@shared/types.js'

export type McpReply = IpcRequest<typeof IpcChannel.McpDocumentReply>

export const refused = (requestId: string, refusal: McpDocumentRefusal): McpReply => ({
  requestId,
  ask: MCP_DOCUMENT_REFUSED,
  refusal,
})
