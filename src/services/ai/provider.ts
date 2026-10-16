export const AiProviderKind = {
  OpenAiCompatible: 'openai-compatible',
  Anthropic: 'anthropic',
  ClaudeCli: 'claude-cli',
  Acp: 'acp',
} as const

export type AiProviderKind = (typeof AiProviderKind)[keyof typeof AiProviderKind]

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

export interface AiPrompt {
  readonly model: string
  readonly system: string
  /** Oldest first, ending with the person's message. */
  readonly messages: readonly AiTurn[]
  readonly maxTokens: number
  readonly fallback?: boolean
}

export type AiStreamPart =
  { readonly text: string } | { readonly costUsd: number } | { readonly model: string }

export interface AiProvider {
  listModels(signal: AbortSignal): Promise<AiModel[]>
  stream(prompt: AiPrompt, signal: AbortSignal): AsyncIterable<AiStreamPart>
}
