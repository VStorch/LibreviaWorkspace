import { AiAction } from '@shared/types.js'
import { Language } from '@shared/i18n/language.js'

export const AI_MAX_OUTPUT_TOKENS = 4096

export interface AiRequest {
  readonly action: AiAction
  readonly text: string
  readonly instruction?: string | undefined
  readonly targetLanguage?: Language | undefined
}

export interface AiMessages {
  readonly system: string
  readonly user: string
}

const LANGUAGE_IN_PROMPT: Readonly<Record<Language, string>> = {
  [Language.Portuguese]: 'Brazilian Portuguese',
  [Language.English]: 'English',
}

const BASE_RULES = [
  'You are a writing assistant inside a word processor.',
  'The user message holds a passage the person selected in their document.',
  'Reply with the resulting text only: no preamble, no explanation, no quotation marks, no Markdown.',
  'Keep paragraph breaks as plain line breaks.',
].join(' ')

const KEEP_LANGUAGE = 'Answer in the language of the passage.'

const ACTION_RULES: Readonly<Record<Exclude<AiAction, typeof AiAction.Translate>, string>> = {
  [AiAction.Rewrite]: `Rewrite the passage so it reads more clearly, keeping its meaning and tone. ${KEEP_LANGUAGE}`,
  [AiAction.Summarize]: `Summarize the passage in a few sentences. ${KEEP_LANGUAGE}`,
  [AiAction.Proofread]: `Fix spelling, grammar and punctuation. Change nothing else. ${KEEP_LANGUAGE}`,
  [AiAction.Custom]: `Follow the instruction that comes before the passage. ${KEEP_LANGUAGE}`,
}

function actionRule(request: AiRequest): string {
  if (request.action !== AiAction.Translate) return ACTION_RULES[request.action]
  const target = LANGUAGE_IN_PROMPT[request.targetLanguage ?? Language.English]
  return `Translate the passage into ${target}, keeping its meaning and tone.`
}

function userMessage(request: AiRequest): string {
  const instruction = request.instruction?.trim() ?? ''
  if (request.action !== AiAction.Custom || instruction === '') return request.text
  return `Instruction: ${instruction}\n\nPassage:\n${request.text}`
}

export function buildMessages(request: AiRequest): AiMessages {
  return { system: `${BASE_RULES} ${actionRule(request)}`, user: userMessage(request) }
}
