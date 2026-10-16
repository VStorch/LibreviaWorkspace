import { z } from 'zod'
import {
  MAX_AI_COMMENT_LENGTH,
  MAX_AI_FIND_LENGTH,
  MAX_AI_PARAGRAPHS,
  MAX_AI_PROPOSALS,
  MAX_AI_READ_PARAGRAPHS,
  MAX_AI_REASON_LENGTH,
  MAX_AI_REPLACE_LENGTH,
  MAX_AI_SEARCH_QUERY_LENGTH,
  MAX_AI_SEARCH_RESULTS,
} from '@shared/limits.js'
import { AiProposalKind, occurrences, type AiProposal } from './proposals.js'
import type { AiToolCall, AiToolDefinition } from './provider.js'

export interface ToolOutcome {
  readonly content: string
  readonly isError: boolean
  readonly proposal?: AiProposal
}

/** The document as the renderer sent it: one text per paragraph, deletions left out. */
export interface ToolContext {
  readonly paragraphs: readonly string[]
  readonly proposalCount: number
}

interface DocumentTool {
  readonly definition: AiToolDefinition
  readonly run: (input: unknown, context: ToolContext) => ToolOutcome
}

export const AiToolName = {
  ReadParagraphs: 'read_paragraphs',
  SearchDocument: 'search_document',
  ProposeEdit: 'propose_edit',
  AddComment: 'add_comment',
} as const

const DEFAULT_SEARCH_RESULTS = 10

const failure = (content: string): ToolOutcome => ({ content, isError: true })
const success = (content: string): ToolOutcome => ({ content, isError: false })

export function numberedParagraph(index: number, text: string): string {
  return `[${index + 1}] ${text}`
}

function parametersOf(schema: z.ZodType): Readonly<Record<string, unknown>> {
  const parameters = { ...(z.toJSONSchema(schema) as Record<string, unknown>) }
  delete parameters['$schema']
  return parameters
}

function defineTool<S extends z.ZodType>(
  name: string,
  description: string,
  schema: S,
  run: (args: z.infer<S>, context: ToolContext) => ToolOutcome,
): DocumentTool {
  return {
    definition: { name, description, parameters: parametersOf(schema) },
    run(input, context) {
      const parsed = schema.safeParse(input)
      if (!parsed.success) return failure(`Invalid arguments: ${z.prettifyError(parsed.error)}`)
      return run(parsed.data, context)
    },
  }
}

const paragraphNumber = z.int().min(1).max(MAX_AI_PARAGRAPHS)
const findText = z
  .string()
  .min(1)
  .max(MAX_AI_FIND_LENGTH)
  .describe('Exact text copied from that paragraph; it must appear there exactly once.')

/** The anchor of a proposal: the paragraph exists and holds `find` exactly once. */
function anchorProblem(paragraph: number, find: string, context: ToolContext): string | null {
  const text = context.paragraphs[paragraph - 1]
  if (text === undefined) {
    return `There is no paragraph ${paragraph}; the document has ${context.paragraphs.length}.`
  }
  const count = occurrences(text, find)
  if (count === 0) return `The text to find is not in paragraph ${paragraph}. It reads: ${text}`
  if (count > 1) {
    return `The text to find appears ${count} times in paragraph ${paragraph}; include more of the words around it.`
  }
  return null
}

function limitProblem(context: ToolContext): string | null {
  return context.proposalCount >= MAX_AI_PROPOSALS
    ? `The limit of ${MAX_AI_PROPOSALS} proposals per message was reached: stop proposing and tell the person.`
    : null
}

const readParagraphs = defineTool(
  AiToolName.ReadParagraphs,
  'Read paragraphs of the document, numbered as [n].',
  z.object({
    from: paragraphNumber.describe('Number of the first paragraph to read.'),
    count: z.int().min(1).max(MAX_AI_READ_PARAGRAPHS),
  }),
  ({ from, count }, { paragraphs }) => {
    if (from > paragraphs.length) return failure(`The document has ${paragraphs.length} paragraphs.`)
    const lines = paragraphs
      .slice(from - 1, from - 1 + count)
      .map((text, offset) => numberedParagraph(from - 1 + offset, text))
    return success(lines.join('\n'))
  },
)

const searchDocument = defineTool(
  AiToolName.SearchDocument,
  'Find the paragraphs that contain a text, ignoring case.',
  z.object({
    query: z.string().min(1).max(MAX_AI_SEARCH_QUERY_LENGTH),
    max: z.int().min(1).max(MAX_AI_SEARCH_RESULTS).optional(),
  }),
  ({ query, max = DEFAULT_SEARCH_RESULTS }, { paragraphs }) => {
    const wanted = query.toLocaleLowerCase()
    const found: string[] = []
    for (let index = 0; index < paragraphs.length && found.length < max; index++) {
      const text = paragraphs[index]!
      if (text.toLocaleLowerCase().includes(wanted)) found.push(numberedParagraph(index, text))
    }
    return success(found.length === 0 ? 'No paragraph contains that text.' : found.join('\n'))
  },
)

const proposeEdit = defineTool(
  AiToolName.ProposeEdit,
  'Propose replacing a text inside one paragraph. The person sees the proposal and decides; nothing changes until then.',
  z.object({
    paragraph: paragraphNumber,
    find: findText,
    replace: z.string().max(MAX_AI_REPLACE_LENGTH).describe('The new text; empty to delete.'),
    reason: z.string().max(MAX_AI_REASON_LENGTH).optional(),
  }),
  ({ paragraph, find, replace, reason }, context) => {
    const problem = limitProblem(context) ?? anchorProblem(paragraph, find, context)
    if (problem !== null) return failure(problem)
    if (find === replace) return failure('The new text is the same as the text to find.')
    const id = String(context.proposalCount + 1)
    return {
      ...success(`Edit ${id} proposed; the person will decide.`),
      proposal: {
        id,
        kind: AiProposalKind.Edit,
        paragraph,
        paragraphText: context.paragraphs[paragraph - 1]!,
        find,
        replace,
        ...(reason === undefined || reason === '' ? {} : { reason }),
      },
    }
  },
)

const addComment = defineTool(
  AiToolName.AddComment,
  'Propose a comment on a text inside one paragraph. The person sees the proposal and decides; nothing changes until then.',
  z.object({
    paragraph: paragraphNumber,
    find: findText,
    comment: z.string().min(1).max(MAX_AI_COMMENT_LENGTH),
  }),
  ({ paragraph, find, comment }, context) => {
    const problem = limitProblem(context) ?? anchorProblem(paragraph, find, context)
    if (problem !== null) return failure(problem)
    const id = String(context.proposalCount + 1)
    return {
      ...success(`Comment ${id} proposed; the person will decide.`),
      proposal: {
        id,
        kind: AiProposalKind.Comment,
        paragraph,
        paragraphText: context.paragraphs[paragraph - 1]!,
        find,
        comment,
      },
    }
  },
)

const TOOLS: ReadonlyMap<string, DocumentTool> = new Map(
  [readParagraphs, searchDocument, proposeEdit, addComment].map((tool) => [tool.definition.name, tool]),
)

export const DOCUMENT_TOOLS: readonly AiToolDefinition[] = [...TOOLS.values()].map((tool) => tool.definition)

function parseArguments(text: string): { readonly value: unknown } | null {
  try {
    return { value: JSON.parse(text.trim() === '' ? '{}' : text) as unknown }
  } catch {
    return null
  }
}

/** Reads only the snapshot: no tool touches the live document, a file, the network or a process. */
export function runDocumentTool(call: AiToolCall, context: ToolContext): ToolOutcome {
  const tool = TOOLS.get(call.name)
  if (tool === undefined) {
    return failure(`There is no tool named "${call.name}". The tools are: ${[...TOOLS.keys()].join(', ')}.`)
  }
  const input = parseArguments(call.arguments)
  if (input === null) return failure('The arguments are not valid JSON.')
  return tool.run(input.value, context)
}
