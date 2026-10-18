import { z } from 'zod'
import type { AiToolCall, AiToolDefinition } from './provider.js'

/** Newest first: a client asking for another gets the newest, as the specification says. */
export const MCP_PROTOCOL_VERSIONS = ['2025-06-18', '2025-03-26', '2024-11-05'] as const

export const JsonRpcErrorCode = {
  ParseError: -32700,
  InvalidRequest: -32600,
  MethodNotFound: -32601,
  InvalidParams: -32602,
  Internal: -32603,
} as const

export const MCP_SERVER_NAME = 'librevia'

export interface McpClient {
  readonly name: string
}

export interface McpToolResult {
  readonly content: string
  readonly isError: boolean
}

export interface McpServerOptions {
  readonly version: string
  readonly tools: readonly AiToolDefinition[]
  readonly runTool: (call: AiToolCall, client: McpClient) => Promise<McpToolResult>
  readonly onInitialize?: (client: McpClient) => void
}

/** One per connection: it remembers the client that introduced itself. */
export interface McpConnection {
  /** The reply line, or null when the message expects none. */
  handle(line: string): Promise<string | null>
}

type RequestId = string | number

type Answer =
  { readonly result: unknown } | { readonly error: { readonly code: number; readonly message: string } }

const requestSchema = z.object({
  jsonrpc: z.literal('2.0'),
  id: z.union([z.string(), z.int()]).optional(),
  method: z.string(),
  params: z.unknown().optional(),
})

const initializeSchema = z.looseObject({
  protocolVersion: z.string().optional(),
  clientInfo: z.looseObject({ name: z.string() }).optional(),
})

const toolCallSchema = z.object({
  name: z.string().min(1),
  arguments: z.record(z.string(), z.unknown()).optional(),
})

const failed = (code: number, message: string): Answer => ({ error: { code, message } })

function line(id: RequestId | null, answer: Answer): string {
  return JSON.stringify({ jsonrpc: '2.0', id, ...answer })
}

export function mcpErrorLine(code: number, message: string): string {
  return line(null, failed(code, message))
}

function parseJson(text: string): { readonly value: unknown } | null {
  try {
    return { value: JSON.parse(text) as unknown }
  } catch {
    return null
  }
}

function negotiatedVersion(requested: string | undefined): string {
  const known = MCP_PROTOCOL_VERSIONS.find((version) => version === requested)
  return known ?? MCP_PROTOCOL_VERSIONS[0]
}

function invalidParams(error: z.ZodError): Answer {
  return failed(JsonRpcErrorCode.InvalidParams, `Invalid params: ${z.prettifyError(error)}`)
}

function toolsList(tools: readonly AiToolDefinition[]): Answer {
  return {
    result: {
      tools: tools.map((tool) => ({
        name: tool.name,
        description: tool.description,
        inputSchema: tool.parameters,
      })),
    },
  }
}

const toolResult = ({ content, isError }: McpToolResult): Answer => ({
  result: { content: [{ type: 'text', text: content }], isError },
})

const TOOL_FAILED: McpToolResult = { content: 'The tool failed; try again.', isError: true }

export function createMcpConnection(options: McpServerOptions): McpConnection {
  let client: McpClient = { name: '' }

  const initialize = (params: unknown): Answer => {
    const parsed = initializeSchema.safeParse(params ?? {})
    if (!parsed.success) return invalidParams(parsed.error)
    client = { name: parsed.data.clientInfo?.name ?? '' }
    options.onInitialize?.(client)
    return {
      result: {
        protocolVersion: negotiatedVersion(parsed.data.protocolVersion),
        capabilities: { tools: {} },
        serverInfo: { name: MCP_SERVER_NAME, version: options.version },
      },
    }
  }

  const callTool = async (params: unknown, id: RequestId): Promise<Answer> => {
    const parsed = toolCallSchema.safeParse(params ?? {})
    if (!parsed.success) return invalidParams(parsed.error)
    const call: AiToolCall = {
      id: String(id),
      name: parsed.data.name,
      arguments: JSON.stringify(parsed.data.arguments ?? {}),
    }
    return toolResult(await options.runTool(call, client).catch(() => TOOL_FAILED))
  }

  const dispatch = async (method: string, params: unknown, id: RequestId): Promise<Answer> => {
    if (method === 'initialize') return initialize(params)
    if (method === 'ping') return { result: {} }
    if (method === 'tools/list') return toolsList(options.tools)
    if (method === 'tools/call') return callTool(params, id)
    return failed(JsonRpcErrorCode.MethodNotFound, `Method not found: ${method}`)
  }

  return {
    async handle(text) {
      const parsed = parseJson(text)
      if (parsed === null) return mcpErrorLine(JsonRpcErrorCode.ParseError, 'Parse error')
      if (Array.isArray(parsed.value)) {
        return mcpErrorLine(JsonRpcErrorCode.InvalidRequest, 'Batches are not supported')
      }
      const request = requestSchema.safeParse(parsed.value)
      if (!request.success) return isReply(parsed.value) ? null : invalidRequest()
      const { id, method, params } = request.data
      if (id === undefined) return null
      return line(id, await dispatch(method, params, id))
    },
  }
}

/** A client's answer to a request of ours: the server sends none, so there is nothing to match. */
function isReply(value: unknown): boolean {
  return typeof value === 'object' && value !== null && !('method' in value) && 'id' in value
}

const invalidRequest = (): string => mcpErrorLine(JsonRpcErrorCode.InvalidRequest, 'Invalid request')
