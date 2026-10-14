export const AiProviderKind = {
  OpenAiCompatible: 'openai-compatible',
  Anthropic: 'anthropic',
  ClaudeCli: 'claude-cli',
} as const

export type AiProviderKind = (typeof AiProviderKind)[keyof typeof AiProviderKind]

export interface AiModel {
  readonly id: string
  readonly label: string
}

export interface AiPrompt {
  readonly model: string
  readonly system: string
  readonly user: string
  readonly maxTokens: number
}

export type AiStreamPart =
  { readonly text: string } | { readonly costUsd: number } | { readonly model: string }

export interface AiProvider {
  listModels(signal: AbortSignal): Promise<AiModel[]>
  stream(prompt: AiPrompt, signal: AbortSignal): AsyncIterable<AiStreamPart>
}
