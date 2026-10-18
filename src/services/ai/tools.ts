import { z } from 'zod'
import type { AiToolCall, AiToolDefinition } from './provider.js'

export interface ToolOutcome<P> {
  readonly content: string
  readonly isError: boolean
  readonly proposal?: P
}

export interface Tool<C, P> {
  readonly definition: AiToolDefinition
  readonly run: (input: unknown, context: C) => ToolOutcome<P>
}

export const failure = (content: string): ToolOutcome<never> => ({ content, isError: true })
export const success = (content: string): ToolOutcome<never> => ({ content, isError: false })

function parametersOf(schema: z.ZodType): Readonly<Record<string, unknown>> {
  const parameters = { ...(z.toJSONSchema(schema) as Record<string, unknown>) }
  delete parameters['$schema']
  return parameters
}

export function defineTool<S extends z.ZodType, C, P>(
  name: string,
  description: string,
  schema: S,
  run: (args: z.infer<S>, context: C) => ToolOutcome<P>,
): Tool<C, P> {
  return {
    definition: { name, description, parameters: parametersOf(schema) },
    run(input, context) {
      const parsed = schema.safeParse(input)
      if (!parsed.success) return failure(`Invalid arguments: ${z.prettifyError(parsed.error)}`)
      return run(parsed.data, context)
    },
  }
}

function parseArguments(text: string): { readonly value: unknown } | null {
  try {
    return { value: JSON.parse(text.trim() === '' ? '{}' : text) as unknown }
  } catch {
    return null
  }
}

export interface Toolbox<C, P> {
  readonly definitions: readonly AiToolDefinition[]
  has(name: string): boolean
  run(call: AiToolCall, context: C): ToolOutcome<P>
}

export function toolbox<C, P>(tools: readonly Tool<C, P>[]): Toolbox<C, P> {
  const byName = new Map(tools.map((tool) => [tool.definition.name, tool]))
  return {
    definitions: tools.map((tool) => tool.definition),
    has: (name) => byName.has(name),
    run(call, context) {
      const tool = byName.get(call.name)
      if (tool === undefined) {
        return failure(
          `There is no tool named "${call.name}". The tools are: ${[...byName.keys()].join(', ')}.`,
        )
      }
      const input = parseArguments(call.arguments)
      if (input === null) return failure('The arguments are not valid JSON.')
      return tool.run(input.value, context)
    },
  }
}
