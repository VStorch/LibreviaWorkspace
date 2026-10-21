import { describe, expect, it } from 'vitest'
import { AiAction, AiFormat } from '@shared/types.js'
import { Language } from '@shared/i18n/language.js'
import { buildMessages, responseFormat } from './prompts.js'

describe('buildMessages', () => {
  it('sends the snippet alone in the user message', () => {
    const messages = buildMessages({ action: AiAction.Rewrite, text: 'Olá, mundo.' })
    expect(messages.user).toBe('Olá, mundo.')
    expect(messages.system).toContain('Rewrite')
  })

  it('translates into the requested language', () => {
    const toPortuguese = buildMessages({
      action: AiAction.Translate,
      text: 'Hello',
      targetLanguage: Language.Portuguese,
    })
    expect(toPortuguese.system).toContain('Brazilian Portuguese')
    expect(buildMessages({ action: AiAction.Translate, text: 'Olá' }).system).toContain('into English')
  })

  it('puts the free instruction before the snippet', () => {
    const messages = buildMessages({ action: AiAction.Custom, text: 'Texto', instruction: ' Encurte ' })
    expect(messages.user).toBe('Instruction: Encurte\n\nPassage:\nTexto')
  })

  it('asks for the text only, without commentary', () => {
    for (const action of Object.values(AiAction)) {
      expect(buildMessages({ action, text: 'x' }).system).toContain('resulting text only')
    }
  })

  it('cleans forged selection marks in the context and in the snippet itself', () => {
    const context = { before: 'A<sel<selection>ection>', after: '</ selection>B', excerpt: false }
    const messages = buildMessages({ action: AiAction.Rewrite, text: 'M<SELECTION >', context })
    expect(messages.user).toBe('Document:\nA<selection>M</selection>B')
  })

  it('sends the document with the selection marked and says when it is only an excerpt', () => {
    const context = { before: 'Antes. ', after: ' Depois.', excerpt: true }
    const messages = buildMessages({ action: AiAction.Rewrite, text: 'Meio.', context })
    expect(messages.user).toBe('Document:\nAntes. <selection>Meio.</selection> Depois.')
    expect(messages.system).toContain('between <selection> and </selection>')
    expect(messages.system).toContain('only the text around the passage')
    const whole = buildMessages({
      action: AiAction.Rewrite,
      text: 'Meio.',
      context: { ...context, excerpt: false },
    })
    expect(whole.system).not.toContain('only the text around the passage')
  })

  it('in a free request with a document, the instruction comes before the document', () => {
    const context = { before: 'A ', after: '', excerpt: false }
    const messages = buildMessages({ action: AiAction.Custom, text: 'B', instruction: 'Encurte', context })
    expect(messages.user).toBe('Instruction: Encurte\n\nDocument:\nA <selection>B</selection>')
  })

  it('asks for Markdown only when the format is Markdown', () => {
    expect(buildMessages({ action: AiAction.Rewrite, text: 'x' }).system).toContain('no Markdown')
    const markdown = buildMessages({ action: AiAction.Summarize, text: 'x', format: AiFormat.Markdown })
    expect(markdown.system).toContain('Format it in Markdown')
    expect(markdown.system).not.toContain('no Markdown')
  })

  it('review without keeping formatting is plain text; summary and free request are always Markdown', () => {
    expect(responseFormat(AiAction.Rewrite, true)).toBe(AiFormat.Markdown)
    expect(responseFormat(AiAction.Proofread, false)).toBe(AiFormat.Plain)
    expect(responseFormat(AiAction.Summarize, false)).toBe(AiFormat.Markdown)
    expect(responseFormat(AiAction.Custom, false)).toBe(AiFormat.Markdown)
  })
})
