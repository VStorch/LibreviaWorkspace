import type { AiModel } from './provider.js'

/** The few fields Librevia reads from the ACP messages; the SDK types fit these shapes. */
export interface AcpUpdate {
  readonly sessionUpdate: string
  readonly content?: unknown
}

export interface AcpPermissionOption {
  readonly optionId: string
  readonly kind: string
}

export type AcpPermissionAnswer = {
  readonly outcome:
    { readonly outcome: 'cancelled' } | { readonly outcome: 'selected'; readonly optionId: string }
}

interface AcpSelectOption {
  readonly value: string
  readonly name: string
}

interface AcpSelectGroup {
  readonly options: readonly AcpSelectOption[]
}

export interface AcpConfigOption {
  readonly id: string
  readonly type: string
  readonly category?: string | null
  readonly currentValue: unknown
  readonly options?: readonly (AcpSelectOption | AcpSelectGroup)[]
}

export interface AcpModelOption {
  readonly configId: string
  readonly current: string
  readonly models: readonly AiModel[]
}

const AUTH_REQUIRED_CODE = -32000
const REJECT_KINDS = ['reject_once', 'reject_always'] as const
const QUOTES = new Set(["'", '"'])

/** Only the answer itself: thoughts, tool calls and plans are the agent's business. */
export function textOfUpdate(update: AcpUpdate): string {
  if (update.sessionUpdate !== 'agent_message_chunk') return ''
  const content = update.content
  if (typeof content !== 'object' || content === null) return ''
  const block = content as { readonly type?: unknown; readonly text?: unknown }
  return block.type === 'text' && typeof block.text === 'string' ? block.text : ''
}

export function refusePermission(options: readonly AcpPermissionOption[]): AcpPermissionAnswer {
  for (const kind of REJECT_KINDS) {
    const option = options.find((candidate) => candidate.kind === kind)
    if (option !== undefined) return { outcome: { outcome: 'selected', optionId: option.optionId } }
  }
  return { outcome: { outcome: 'cancelled' } }
}

export function modelOptionOf(
  configOptions: readonly AcpConfigOption[] | null | undefined,
): AcpModelOption | undefined {
  const option = configOptions?.find(
    (candidate) => candidate.category === 'model' && candidate.type === 'select',
  )
  if (option === undefined || typeof option.currentValue !== 'string') return undefined
  const flat = (option.options ?? []).flatMap((entry) => ('options' in entry ? entry.options : [entry]))
  return {
    configId: option.id,
    current: option.currentValue,
    models: flat.map((entry) => ({ id: entry.value, label: entry.name })),
  }
}

export function modelsOf(configOptions: readonly AcpConfigOption[] | null | undefined): AiModel[] {
  return [...(modelOptionOf(configOptions)?.models ?? [])]
}

export function isAuthError(cause: unknown): boolean {
  return (
    typeof cause === 'object' && cause !== null && (cause as { code?: unknown }).code === AUTH_REQUIRED_CODE
  )
}

/** Spaces separate arguments; single or double quotes keep spaces inside one. No shell, no escapes. */
export function splitArgs(text: string): string[] {
  const args: string[] = []
  let current: string | null = null
  let quote: string | null = null
  for (const char of text) {
    if (quote !== null) {
      if (char === quote) quote = null
      else current = (current ?? '') + char
    } else if (QUOTES.has(char)) {
      quote = char
      current ??= ''
    } else if (/\s/u.test(char)) {
      if (current !== null) args.push(current)
      current = null
    } else current = (current ?? '') + char
  }
  if (current !== null) args.push(current)
  return args
}

export function joinArgs(args: readonly string[]): string {
  return args.map(quoteArg).join(' ')
}

function quoteArg(arg: string): string {
  if (arg !== '' && !/[\s'"]/u.test(arg)) return arg
  return arg.includes("'") ? `"${arg}"` : `'${arg}'`
}
