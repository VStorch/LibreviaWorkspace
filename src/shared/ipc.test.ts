import { describe, expect, it } from 'vitest'
import { INVOCABLE_IPC_CHANNELS, IpcChannel, PUSH_IPC_CHANNELS } from './ipc-channels.js'
import { MAX_TEXT_LENGTH, ipcContracts, pushContracts } from './ipc.js'

describe('contratos de IPC', () => {
  it('define um contrato para cada canal invocável', () => {
    // A channel without a schema would pass unvalidated payloads to the handler.
    for (const channel of INVOCABLE_IPC_CHANNELS) {
      expect(ipcContracts[channel]).toBeDefined()
    }
  })

  it('não expõe contrato para canal não declarado', () => {
    expect(Object.keys(ipcContracts).sort()).toEqual([...INVOCABLE_IPC_CHANNELS].sort())
  })

  it('mantém o canal de menu fora dos invocáveis: ele vai de main para renderer', () => {
    expect(INVOCABLE_IPC_CHANNELS).not.toContain(IpcChannel.MenuCommand)
    expect(ipcContracts).not.toHaveProperty(IpcChannel.MenuCommand)
  })
})

describe('contratos do sentido main → renderer', () => {
  it('define um contrato para cada canal empurrado', () => {
    // Without a schema nobody would check the shape, and the renderer would drop the message
    // silently, the costliest defect here.
    for (const channel of PUSH_IPC_CHANNELS) {
      expect(pushContracts[channel]).toBeDefined()
    }
  })

  it('os dois sentidos não se misturam', () => {
    for (const channel of PUSH_IPC_CHANNELS) {
      expect(INVOCABLE_IPC_CHANNELS).not.toContain(channel)
      expect(ipcContracts).not.toHaveProperty(channel)
    }
  })
})

describe('validação do alvo do menu de contexto', () => {
  const schema = pushContracts[IpcChannel.ContextMenuRequested]

  const alvo = {
    x: 120,
    y: 40,
    editable: true,
    misspelledWord: 'abacaxxi',
    dictionarySuggestions: ['abacaxi'],
    canCut: true,
    canCopy: true,
    canPaste: true,
  }

  it('aceita o que o Chromium manda', () => {
    expect(schema.safeParse(alvo).success).toBe(true)
  })

  it('recusa coordenada negativa e lista de sugestões absurda', () => {
    // This becomes a screen position and menu items: a negative coordinate puts the menu outside
    // the window, and thirty suggestions overflow it.
    expect(schema.safeParse({ ...alvo, x: -1 }).success).toBe(false)
    expect(
      schema.safeParse({ ...alvo, dictionarySuggestions: Array.from({ length: 30 }, () => 'x') }).success,
    ).toBe(false)
  })
})

describe('validação de prefs:set', () => {
  const schema = ipcContracts[IpcChannel.PreferencesSet].request

  it('aceita um remendo de uma chave só', () => {
    // Whoever toggles formatting marks has no opinion about spelling.
    expect(schema.parse({ invisibleCharacters: true })).toEqual({ invisibleCharacters: true })
  })

  it('recusa valor que não é booleano', () => {
    expect(schema.safeParse({ spellcheck: 'sim' }).success).toBe(false)
  })
})

describe('validação de file:save', () => {
  const schema = ipcContracts[IpcChannel.FileSave].request

  it('aceita uma gravação bem formada', () => {
    expect(schema.safeParse({ path: '/home/ana/ata.txt', content: 'texto', origin: null }).success).toBe(true)
  })

  it('aceita a origem do documento aberto', () => {
    const payload = { path: '/home/ana/ata.docx', content: '{}', origin: '/home/ana/ata.sdoc' }
    expect(schema.safeParse(payload).success).toBe(true)
  })

  it.each([
    ['sem caminho', { content: 'texto' }],
    ['caminho vazio', { path: '', content: 'texto' }],
    ['sem conteúdo', { path: '/a/b.txt' }],
    ['caminho não textual', { path: 42, content: 'texto', origin: null }],
    ['sem origem', { path: '/a/b.docx', content: 'texto' }],
    ['origem vazia', { path: '/a/b.docx', content: 'texto', origin: '' }],
    ['nulo', null],
  ])('recusa %s', (_label, payload) => {
    expect(schema.safeParse(payload).success).toBe(false)
  })

  it('recusa conteúdo acima do teto de memória', () => {
    const oversized = { path: '/a/b.txt', content: 'x'.repeat(MAX_TEXT_LENGTH + 1) }
    expect(schema.safeParse(oversized).success).toBe(false)
  })

  it('descarta campos não previstos no contrato', () => {
    const parsed = schema.parse({ path: '/a/b.txt', content: 'oi', origin: null, extra: 'ignorar' })
    expect(parsed).toEqual({ path: '/a/b.txt', content: 'oi', origin: null })
  })
})

describe('modelos do Word (M11)', () => {
  const open = ipcContracts[IpcChannel.TemplateOpen].request
  const opened = ipcContracts[IpcChannel.TemplateOpen].response
  const list = ipcContracts[IpcChannel.TemplateList].response

  it('abre pelo par fonte e id', () => {
    expect(open.safeParse({ source: 'builtin', id: 'carta.dotx' }).success).toBe(true)
    expect(open.safeParse({ source: 'user', id: '/home/ana/Modelos/proposta.dotx' }).success).toBe(true)
  })

  it.each([
    ['fonte desconhecida', { source: 'web', id: 'carta.dotx' }],
    ['id vazio', { source: 'builtin', id: '' }],
    ['sem id', { source: 'user' }],
  ])('recusa %s', (_label, payload) => {
    expect(open.safeParse(payload).success).toBe(false)
  })

  it('o documento aberto de um modelo traz a marca de modelo', () => {
    const file = {
      path: '/r/carta.dotx',
      name: 'Carta.docx',
      kind: 'document',
      content: '{}',
      template: true,
    }
    expect(opened.parse({ file }).file.template).toBe(true)
  })

  it('a galeria lista as duas fontes e a pasta do usuário', () => {
    const entry = { source: 'builtin', id: 'carta.dotx', name: 'Carta', description: 'Uma carta' }
    expect(list.safeParse({ builtin: [entry], user: [], folder: '/home/ana/Modelos' }).success).toBe(true)
    expect(list.safeParse({ builtin: [{ ...entry, source: 'outra' }], user: [], folder: '' }).success).toBe(
      false,
    )
  })
})

describe('validação de file:open-recent', () => {
  const schema = ipcContracts[IpcChannel.FileOpenRecent].request

  it('exige um caminho não vazio', () => {
    expect(schema.safeParse({ path: '' }).success).toBe(false)
    expect(schema.safeParse({ path: '/a/b.txt' }).success).toBe(true)
  })
})

describe('validação de fonts:list', () => {
  // `src/main/ipc/registry.ts` runs this schema on the real response, and `registry.test.ts` proves
  // it: a schema nobody runs is a compile-time type, not protection.
  const schema = ipcContracts[IpcChannel.FontsList].response

  it('aceita a lista vazia', () => {
    // A system without `fontconfig` returns nothing, and that is not an error: the toolbar keeps
    // the bundled fonts.
    expect(schema.safeParse({ families: [] }).success).toBe(true)
  })

  it('recusa nome vazio e lista absurda', () => {
    // This is the output of a system program, data and not truth: an empty name would become an
    // invisible option, and a list of a hundred thousand entries would freeze the toolbar.
    expect(schema.safeParse({ families: [''] }).success).toBe(false)
    expect(schema.safeParse({ families: Array.from({ length: 4001 }, () => 'Arial') }).success).toBe(false)
  })
})

describe('validação de dialog:confirm-discard', () => {
  const response = ipcContracts[IpcChannel.DialogConfirmDiscard].response

  it.each(['save', 'discard', 'cancel'])('aceita a resposta %s', (choice) => {
    expect(response.safeParse({ choice }).success).toBe(true)
  })

  it('recusa uma resposta fora das três previstas', () => {
    expect(response.safeParse({ choice: 'talvez' }).success).toBe(false)
  })
})
