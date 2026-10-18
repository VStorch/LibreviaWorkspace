import type { AiModel, AiToolServer } from './provider.js'

/** The few fields Librevia reads from the ACP messages; the SDK types fit these shapes. */
export interface AcpUpdate {
  readonly sessionUpdate: string
  readonly content?: unknown
}

export interface AcpPermissionOption {
  readonly optionId: string
  readonly kind: string
}

/** What the agent says about the tool it wants to run; any field may be missing. */
export interface AcpToolCall {
  readonly name?: string | null
  readonly title?: string | null
  readonly kind?: string | null
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
/** The kinds an agent gives its own tools that change files, run programs or reach the network. */
const UNSAFE_TOOL_KINDS = new Set(['edit', 'delete', 'move', 'execute', 'fetch'])
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

/**
 * How agents name a tool of an MCP server: bare, as Claude Code (`mcp__librevia__read_paragraphs`),
 * as Gemini (`mcp_librevia_read_paragraphs`, or `read_paragraphs (librevia MCP Server)` and
 * `read_paragraphs(from: 1)` as titles).
 */
function namesOf(server: string, tool: string): string[] {
  return [tool, `mcp__${server}__${tool}`, `mcp_${server}_${tool}`, `${tool} (${server} MCP Server)`]
}

function isTitleOf(title: string, server: string, tool: string): boolean {
  return namesOf(server, tool).includes(title) || title.startsWith(`${tool}(`)
}

/** Only an exact name of one of Librevia's tools, never of a tool that edits, runs or fetches. */
export function isLibreviaTool(call: AcpToolCall, server: AiToolServer | undefined): boolean {
  if (server === undefined || UNSAFE_TOOL_KINDS.has(call.kind ?? '')) return false
  const { name, title } = call
  return server.tools.some((tool) => {
    if (typeof name === 'string' && name !== '') return namesOf(server.name, tool).includes(name)
    return typeof title === 'string' && isTitleOf(title.trim(), server.name, tool)
  })
}

/** Librevia's tools run once each time they are asked for; everything else is refused. */
export function answerPermission(
  call: AcpToolCall,
  options: readonly AcpPermissionOption[],
  server: AiToolServer | undefined,
): AcpPermissionAnswer {
  const allow = options.find((option) => option.kind === 'allow_once')
  if (allow === undefined || !isLibreviaTool(call, server)) return refusePermission(options)
  return { outcome: { outcome: 'selected', optionId: allow.optionId } }
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
