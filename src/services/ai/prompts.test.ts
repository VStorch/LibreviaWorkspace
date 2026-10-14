import { describe, expect, it } from 'vitest'
import { AiAction } from '@shared/types.js'
import { Language } from '@shared/i18n/language.js'
import { buildMessages } from './prompts.js'

describe('buildMessages', () => {
  it('manda o trecho sozinho na mensagem do usuário', () => {
    const messages = buildMessages({ action: AiAction.Rewrite, text: 'Olá, mundo.' })
    expect(messages.user).toBe('Olá, mundo.')
    expect(messages.system).toContain('Rewrite')
  })

  it('traduz para o idioma pedido', () => {
    const toPortuguese = buildMessages({
      action: AiAction.Translate,
      text: 'Hello',
      targetLanguage: Language.Portuguese,
    })
    expect(toPortuguese.system).toContain('Brazilian Portuguese')
    expect(buildMessages({ action: AiAction.Translate, text: 'Olá' }).system).toContain('into English')
  })

  it('põe a instrução livre antes do trecho', () => {
    const messages = buildMessages({ action: AiAction.Custom, text: 'Texto', instruction: ' Encurte ' })
    expect(messages.user).toBe('Instruction: Encurte\n\nPassage:\nTexto')
  })

  it('pede só o texto, sem comentário', () => {
    for (const action of Object.values(AiAction)) {
      expect(buildMessages({ action, text: 'x' }).system).toContain('resulting text only')
    }
  })
})
