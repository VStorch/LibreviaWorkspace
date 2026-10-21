import { describe, expect, it } from 'vitest'
import { ErrorCode } from '@shared/errors.js'
import { DEFAULT_LOCAL_AI_CONTEXT_LENGTH, MAX_AI_DOCUMENT_LENGTH } from '@shared/limits.js'
import { AiSettings, contextLengthOf, type AiSettingsData } from './settings.js'

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

describe('assistants settings', () => {
  it('never returns the key in the list, only that it exists', () => {
    const { subject, stored } = settings()
    const list = subject.save({ ...draft, apiKey: 'sk-secreta' })
    expect(JSON.stringify(list)).not.toContain('sk-secreta')
    expect(list.providers[0]).toMatchObject({ id: 'id1', hasKey: true, local: false, consented: false })
    expect(stored().providers[0]?.sealedKey).toBe('selado:sk-secreta')
    expect(subject.keyOf(subject.find('id1'))).toBe('sk-secreta')
  })

  it('the first becomes the default, and removing the default passes to the next', () => {
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

  it('keeps the key when it is not sent, and forgets the consent when the address changes', () => {
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

  it('discards the old key when the destination changes, so it does not go to another server', () => {
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

  it('revokes the consent on request', () => {
    const { subject } = settings()
    subject.save({ ...draft, apiKey: 'k' })
    subject.markConsented('id1')
    expect(subject.save({ ...draft, id: 'id1', revokeConsent: true }).providers[0]?.consented).toBe(false)
  })

  it('requires an address for an OpenAI-compatible server and flags a nonexistent provider', () => {
    const { subject } = settings()
    expect(() => subject.save({ ...draft, kind: 'openai-compatible' })).toThrow()
    expect(() => subject.find('nenhum')).toThrow(expect.objectContaining({ code: ErrorCode.AiNotConfigured }))
  })

  it('an old file, without preferences, stays valid and gets the defaults', () => {
    let data: unknown = {
      providers: [{ id: 'a', kind: 'anthropic', label: 'Antigo', consented: true }],
      defaultProviderId: 'a',
    }
    const subject = new AiSettings(
      { read: () => data, write: (value) => (data = value) },
      { seal: (plain) => plain, open: (sealed) => sealed },
    )
    expect(subject.list()).toMatchObject({
      providers: [{ id: 'a', label: 'Antigo' }],
      preferences: { keepFormatting: true, anthropicFallback: false, chatProposals: false },
    })
  })

  it('preferences saved before proposals and the MCP server get the checkboxes off', () => {
    const data: unknown = {
      providers: [],
      defaultProviderId: null,
      preferences: { keepFormatting: false, anthropicFallback: true },
    }
    const subject = new AiSettings(
      { read: () => data, write: () => undefined },
      { seal: (plain) => plain, open: (sealed) => sealed },
    )
    expect(subject.preferences()).toEqual({
      keepFormatting: false,
      anthropicFallback: true,
      chatProposals: false,
      mcpServer: false,
    })
  })

  it('consent given when only the selection was sent is asked again, now that the document goes', () => {
    let data: unknown = {
      providers: [{ id: 'a', kind: 'anthropic', label: 'Antigo', consented: true }],
      defaultProviderId: 'a',
    }
    const subject = new AiSettings(
      { read: () => data, write: (value) => (data = value) },
      { seal: (plain) => plain, open: (sealed) => sealed },
    )
    expect(subject.list().providers[0]?.consented).toBe(false)
    subject.markConsented('a')
    expect(subject.list().providers[0]?.consented).toBe(true)
  })

  it('the context limit applies only to an OpenAI-compatible server, and local starts smaller', () => {
    const { subject } = settings()
    const local = {
      kind: 'openai-compatible',
      label: 'Ollama',
      baseUrl: 'http://127.0.0.1:11434/v1',
      makeDefault: false,
      revokeConsent: false,
    } as const
    subject.save(local)
    expect(contextLengthOf(subject.find('id1'))).toBe(DEFAULT_LOCAL_AI_CONTEXT_LENGTH)
    subject.save({ ...local, id: 'id1', contextLength: 12_000 })
    expect(subject.list().providers[0]?.contextLength).toBe(12_000)
    expect(contextLengthOf(subject.find('id1'))).toBe(12_000)
    subject.save({ ...draft, contextLength: 10 })
    expect(contextLengthOf(subject.find('id2'))).toBe(MAX_AI_DOCUMENT_LENGTH)
  })

  it('choosing the model does not touch the key or the consent, and preferences survive', () => {
    const { subject, stored } = settings()
    subject.save({ ...draft, apiKey: 'k' })
    subject.markConsented('id1')
    subject.savePreferences({
      keepFormatting: false,
      anthropicFallback: true,
      chatProposals: true,
      mcpServer: false,
    })
    expect(subject.chooseModel('id1', 'modelo-x').providers[0]).toMatchObject({
      defaultModel: 'modelo-x',
      hasKey: true,
      consented: true,
    })
    expect(subject.chooseModel('id1', '').providers[0]?.defaultModel).toBeUndefined()
    subject.save({ ...draft, id: 'id1', label: 'Renomeado' })
    expect(stored().preferences).toEqual({
      keepFormatting: false,
      anthropicFallback: true,
      chatProposals: true,
      mcpServer: false,
    })
    expect(stored().providers[0]).toMatchObject({
      label: 'Renomeado',
      sealedKey: 'selado:k',
      consented: true,
    })
  })

  it('a known ACP agent stores the id and the table arguments, never a path coming from the renderer', () => {
    const { subject, stored } = settings()
    const list = subject.save({
      kind: 'acp',
      label: 'Gemini',
      agentId: 'gemini',
      makeDefault: false,
      revokeConsent: false,
    })
    expect(stored().providers[0]).toMatchObject({ agentId: 'gemini', args: ['--acp'] })
    expect(stored().providers[0]?.command).toBeUndefined()
    expect(list.providers[0]).toMatchObject({ agentId: 'gemini', args: '--acp', local: false })
  })

  it('changing an agent arguments forgets the consent', () => {
    const { subject } = settings()
    const gemini = {
      kind: 'acp',
      label: 'Gemini',
      agentId: 'gemini',
      makeDefault: false,
      revokeConsent: false,
    } as const
    subject.save(gemini)
    subject.markConsented('id1')
    expect(subject.save({ ...gemini, id: 'id1', args: '--acp' }).providers[0]?.consented).toBe(true)
    expect(subject.save({ ...gemini, id: 'id1', args: '--acp --outro' }).providers[0]?.consented).toBe(false)
  })

  it('an ACP agent chosen in the dialog shows only the file name; renaming keeps the consent and changing the program forgets it', () => {
    const { subject, stored } = settings()
    subject.saveAgentExecutable(
      { label: 'Meu agente', args: "--acp --dir 'a b'" },
      '/opt/agente/bin/meu-agente',
    )
    expect(stored().providers[0]).toMatchObject({
      command: '/opt/agente/bin/meu-agente',
      args: ['--acp', '--dir', 'a b'],
    })
    expect(JSON.stringify(subject.list())).not.toContain('/opt/agente')
    expect(subject.list().providers[0]).toMatchObject({
      commandName: 'meu-agente',
      args: "--acp --dir 'a b'",
    })

    subject.markConsented('id1')
    subject.save({
      id: 'id1',
      kind: 'acp',
      label: 'Renomeado',
      args: "--acp --dir 'a b'",
      makeDefault: false,
      revokeConsent: false,
    })
    expect(stored().providers[0]).toMatchObject({ command: '/opt/agente/bin/meu-agente', label: 'Renomeado' })
    expect(subject.list().providers[0]?.consented).toBe(true)

    subject.saveAgentExecutable({ id: 'id1', label: 'Renomeado', args: '' }, '/usr/bin/outro')
    expect(subject.list().providers[0]?.consented).toBe(false)
  })

  it('a new custom ACP agent without a chosen executable is refused', () => {
    const { subject } = settings()
    expect(() =>
      subject.save({ kind: 'acp', label: 'X', makeDefault: false, revokeConsent: false }),
    ).toThrow()
  })
})
