export const AiProviderKind = {
  OpenAiCompatible: 'openai-compatible',
  Anthropic: 'anthropic',
  ClaudeCli: 'claude-cli',
  Acp: 'acp',
} as const

export type AiProviderKind = (typeof AiProviderKind)[keyof typeof AiProviderKind]

/**
 * Claude Code and the ACP agents run the tool loop themselves and reach the document's tools
 * through Librevia's local MCP server, with a key for one message; the others get them here.
 */
export function runsToolsThroughMcp(kind: AiProviderKind): boolean {
  return kind === AiProviderKind.ClaudeCli || kind === AiProviderKind.Acp
}

export interface AiModel {
  readonly id: string
  readonly label: string
  /** Whether the provider may hand the request to another model when this one refuses. */
  readonly fallback?: boolean
}

export const AiRole = { User: 'user', Assistant: 'assistant' } as const

export type AiRole = (typeof AiRole)[keyof typeof AiRole]

export interface AiTurn {
  readonly role: AiRole
  readonly text: string
}

export interface AiToolDefinition {
  readonly name: string
  readonly description: string
  /** JSON Schema of the arguments object. */
  readonly parameters: Readonly<Record<string, unknown>>
}

export interface AiToolCall {
  readonly id: string
  readonly name: string
  /** As the model wrote it: JSON that may not parse. */
  readonly arguments: string
}

export interface AiToolResult {
  readonly callId: string
  readonly content: string
  readonly isError: boolean
}

/** One round inside the current reply: what the model said, the tools it called and what they gave back. */
export interface AiToolExchange {
  readonly text: string
  readonly calls: readonly AiToolCall[]
  readonly results: readonly AiToolResult[]
}

/** The local MCP server with the document's tools, for a provider that runs the tool loop itself. */
export interface AiToolServer {
  readonly name: string
  readonly command: string
  readonly args: readonly string[]
  readonly env: Readonly<Record<string, string>>
  readonly tools: readonly string[]
}

export interface AiPrompt {
  readonly model: string
  readonly system: string
  /** Oldest first, ending with the person's message. */
  readonly messages: readonly AiTurn[]
  readonly maxTokens: number
  readonly fallback?: boolean
  readonly tools?: readonly AiToolDefinition[]
  /** After the person's last message, oldest first. */
  readonly exchanges?: readonly AiToolExchange[]
  readonly toolServer?: AiToolServer
}

export type AiStreamPart =
  | { readonly text: string }
  | { readonly costUsd: number }
  | { readonly model: string }
  | { readonly toolCall: AiToolCall }

export interface AiProvider {
  listModels(signal: AbortSignal): Promise<AiModel[]>
  stream(prompt: AiPrompt, signal: AbortSignal): AsyncIterable<AiStreamPart>
}
