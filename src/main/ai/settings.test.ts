import { describe, expect, it } from 'vitest'
import { ErrorCode } from '@shared/errors.js'
import { AiSettings, type AiSettingsData } from './settings.js'

function settings(): { subject: AiSettings; stored: () => AiSettingsData } {
  let data: unknown = undefined
  let next = 0
  const subject = new AiSettings(
    { read: () => data, write: (value) => (data = value) },
    { seal: (plain) => `selado:${plain}`, open: (sealed) => sealed.replace('selado:', '') },
    () => `id${++next}`,
  )
  return { subject, stored: () => data as AiSettingsData }
}

const draft = { kind: 'anthropic', label: 'Anthropic', makeDefault: false, revokeConsent: false } as const

describe('configuração dos assistentes', () => {
  it('nunca devolve a chave na lista, só que ela existe', () => {
    const { subject, stored } = settings()
    const list = subject.save({ ...draft, apiKey: 'sk-secreta' })
    expect(JSON.stringify(list)).not.toContain('sk-secreta')
    expect(list.providers[0]).toMatchObject({ id: 'id1', hasKey: true, local: false, consented: false })
    expect(stored().providers[0]?.sealedKey).toBe('selado:sk-secreta')
    expect(subject.keyOf(subject.find('id1'))).toBe('sk-secreta')
  })

  it('o primeiro vira padrão, e remover o padrão passa para o seguinte', () => {
    const { subject } = settings()
    subject.save({ ...draft, apiKey: 'a' })
    subject.save({
      kind: 'openai-compatible',
      label: 'Ollama',
      baseUrl: 'http://localhost:11434/v1',
      makeDefault: false,
      revokeConsent: false,
    })
    expect(subject.list().defaultProviderId).toBe('id1')
    expect(subject.list().providers[1]?.local).toBe(true)
    expect(subject.remove('id1').defaultProviderId).toBe('id2')
  })

  it('manter a chave quando ela não vem, e esquecer o consentimento ao mudar o endereço', () => {
    const { subject } = settings()
    const base = {
      kind: 'openai-compatible',
      label: 'Servidor',
      makeDefault: false,
      revokeConsent: false,
    } as const
    subject.save({ ...base, baseUrl: 'https://a.example/v1', apiKey: 'k' })
    subject.markConsented('id1')
    expect(subject.save({ ...base, id: 'id1', baseUrl: 'https://a.example/v1' }).providers[0]).toMatchObject({
      hasKey: true,
      consented: true,
    })
    expect(
      subject.save({ ...base, id: 'id1', baseUrl: 'https://b.example/v1' }).providers[0]?.consented,
    ).toBe(false)
  })

  it('descarta a chave antiga quando o destino muda, para ela não seguir para outro servidor', () => {
    const { subject } = settings()
    const base = {
      kind: 'openai-compatible',
      label: 'Servidor',
      makeDefault: false,
      revokeConsent: false,
    } as const
    subject.save({ ...base, baseUrl: 'https://a.example/v1', apiKey: 'k' })
    subject.markConsented('id1')
    expect(subject.save({ ...base, id: 'id1', baseUrl: 'https://b.example/v1' }).providers[0]).toMatchObject({
      hasKey: false,
      consented: false,
    })
    subject.save({ ...draft, apiKey: 'k2' })
    expect(subject.save({ ...base, id: 'id2', baseUrl: 'https://c.example/v1' }).providers[1]?.hasKey).toBe(
      false,
    )
  })

  it('revoga o consentimento a pedido', () => {
    const { subject } = settings()
    subject.save({ ...draft, apiKey: 'k' })
    subject.markConsented('id1')
    expect(subject.save({ ...draft, id: 'id1', revokeConsent: true }).providers[0]?.consented).toBe(false)
  })

  it('exige endereço para servidor compatível com OpenAI e acusa provedor inexistente', () => {
    const { subject } = settings()
    expect(() => subject.save({ ...draft, kind: 'openai-compatible' })).toThrow()
    expect(() => subject.find('nenhum')).toThrow(expect.objectContaining({ code: ErrorCode.AiNotConfigured }))
  })

  it('arquivo antigo, sem preferências, continua valendo e ganha os padrões', () => {
    let data: unknown = {
      providers: [{ id: 'a', kind: 'anthropic', label: 'Antigo', consented: true }],
      defaultProviderId: 'a',
    }
    const subject = new AiSettings(
      { read: () => data, write: (value) => (data = value) },
      { seal: (plain) => plain, open: (sealed) => sealed },
    )
    expect(subject.list()).toMatchObject({
      providers: [{ id: 'a', label: 'Antigo', consented: true }],
      preferences: { keepFormatting: true, anthropicFallback: false },
    })
  })

  it('escolher o modelo não mexe na chave nem no consentimento, e as preferências sobrevivem', () => {
    const { subject, stored } = settings()
    subject.save({ ...draft, apiKey: 'k' })
    subject.markConsented('id1')
    subject.savePreferences({ keepFormatting: false, anthropicFallback: true })
    expect(subject.chooseModel('id1', 'modelo-x').providers[0]).toMatchObject({
      defaultModel: 'modelo-x',
      hasKey: true,
      consented: true,
    })
    expect(subject.chooseModel('id1', '').providers[0]?.defaultModel).toBeUndefined()
    subject.save({ ...draft, id: 'id1', label: 'Renomeado' })
    expect(stored().preferences).toEqual({ keepFormatting: false, anthropicFallback: true })
    expect(stored().providers[0]).toMatchObject({
      label: 'Renomeado',
      sealedKey: 'selado:k',
      consented: true,
    })
  })
})
