import { AiAction, AiFormat } from '@shared/types.js'
import { Language } from '@shared/i18n/language.js'
import {
  NO_CONTEXT,
  SELECTION_CLOSE,
  SELECTION_OPEN,
  stripSelectionTags,
  type DocumentContext,
} from './context.js'

export const AI_MAX_OUTPUT_TOKENS = 4096

export interface AiRequest {
  readonly action: AiAction
  readonly text: string
  readonly instruction?: string | undefined
  readonly targetLanguage?: Language | undefined
  readonly format?: AiFormat | undefined
  readonly context?: DocumentContext | undefined
}

export interface AiMessages {
  readonly system: string
  readonly user: string
}

const LANGUAGE_IN_PROMPT: Readonly<Record<Language, string>> = {
  [Language.Portuguese]: 'Brazilian Portuguese',
  [Language.English]: 'English',
}

const ROLE = 'You are a writing assistant inside a word processor.'
const PASSAGE_ALONE = 'The user message holds a passage the person selected in their document.'
const PASSAGE_IN_DOCUMENT = [
  `The user message holds the person's document, with the passage they selected between ${SELECTION_OPEN} and ${SELECTION_CLOSE}.`,
  'Work on that passage only; the rest of the document is there for context.',
].join(' ')
const EXCERPT = 'The document is too long to send whole: you see only the text around the passage.'
const ONLY_RESULT = 'Reply with the resulting text only: no preamble, no explanation, no quotation marks'

const FORMAT_RULES: Readonly<Record<AiFormat, string>> = {
  [AiFormat.Plain]: `${ONLY_RESULT}, no Markdown. Keep paragraph breaks as plain line breaks.`,
  [AiFormat.Markdown]: [
    `${ONLY_RESULT}, no selection tags.`,
    'Format it in Markdown (bold, italic, headings, lists, links) where the text calls for it, without HTML.',
    'The passage is in Markdown too: keep its formatting unless asked otherwise.',
  ].join(' '),
}

const KEEP_LANGUAGE = 'Answer in the language of the passage.'

const ACTION_RULES: Readonly<Record<Exclude<AiAction, typeof AiAction.Translate>, string>> = {
  [AiAction.Rewrite]: `Rewrite the passage so it reads more clearly, keeping its meaning and tone. ${KEEP_LANGUAGE}`,
  [AiAction.Summarize]: `Summarize the passage in a few sentences. ${KEEP_LANGUAGE}`,
  [AiAction.Proofread]: `Fix spelling, grammar and punctuation. Change nothing else. ${KEEP_LANGUAGE}`,
  [AiAction.Custom]: `Follow the instruction at the start of the user message. ${KEEP_LANGUAGE}`,
}

function actionRule(request: AiRequest): string {
  if (request.action !== AiAction.Translate) return ACTION_RULES[request.action]
  const target = LANGUAGE_IN_PROMPT[request.targetLanguage ?? Language.English]
  return `Translate the passage into ${target}, keeping its meaning and tone.`
}

function hasContext(context: DocumentContext): boolean {
  return context.before !== '' || context.after !== ''
}

function documentOf(request: AiRequest, context: DocumentContext): string {
  if (!hasContext(context)) return `Passage:\n${request.text}`
  return `Document:\n${stripSelectionTags(context.before)}${SELECTION_OPEN}${stripSelectionTags(request.text)}${SELECTION_CLOSE}${stripSelectionTags(context.after)}`
}

function userMessage(request: AiRequest, context: DocumentContext): string {
  const instruction = request.instruction?.trim() ?? ''
  const withInstruction = request.action === AiAction.Custom && instruction !== ''
  if (!withInstruction && !hasContext(context)) return request.text
  const document = documentOf(request, context)
  return withInstruction ? `Instruction: ${instruction}\n\n${document}` : document
}

function systemMessage(request: AiRequest, context: DocumentContext): string {
  const passage = hasContext(context) ? PASSAGE_IN_DOCUMENT : PASSAGE_ALONE
  return [
    ROLE,
    passage,
    ...(context.excerpt ? [EXCERPT] : []),
    FORMAT_RULES[request.format ?? AiFormat.Plain],
    actionRule(request),
  ].join(' ')
}

/** The actions that revise the passage; the others write new text, free to format it. */
const REVISING_ACTIONS: ReadonlySet<AiAction> = new Set([
  AiAction.Rewrite,
  AiAction.Proofread,
  AiAction.Translate,
])

export function responseFormat(action: AiAction, keepFormatting: boolean): AiFormat {
  return REVISING_ACTIONS.has(action) && !keepFormatting ? AiFormat.Plain : AiFormat.Markdown
}

export function buildMessages(request: AiRequest): AiMessages {
  const context = request.context ?? NO_CONTEXT
  return { system: systemMessage(request, context), user: userMessage(request, context) }
}
