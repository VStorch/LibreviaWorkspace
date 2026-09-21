import { describe, expect, it } from 'vitest'
import { INVOCABLE_IPC_CHANNELS, IpcChannel, PUSH_IPC_CHANNELS } from './ipc-channels.js'
import { MAX_TEXT_LENGTH, ipcContracts, pushContracts } from './ipc.js'

describe('contratos de IPC', () => {
  it('define um contrato para cada canal invocável', () => {
    // Um canal sem schema passaria payload não validada ao handler.
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
    // Sem schema, a mensagem chegaria ao renderer sem ninguém conferir a forma —
    // e o renderer a descartaria em silêncio, que é o defeito mais caro daqui.
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
    // O que chega aqui vira posição na tela e itens de menu: uma coordenada
    // negativa põe o menu fora da janela, e trinta sugestões o fazem sair dela.
    expect(schema.safeParse({ ...alvo, x: -1 }).success).toBe(false)
    expect(
      schema.safeParse({ ...alvo, dictionarySuggestions: Array.from({ length: 30 }, () => 'x') }).success,
    ).toBe(false)
  })
})

describe('validação de prefs:set', () => {
  const schema = ipcContracts[IpcChannel.PreferencesSet].request

  it('aceita um remendo de uma chave só', () => {
    // Quem clica em "marcas de formatação" não tem opinião sobre ortografia.
    expect(schema.parse({ invisibleCharacters: true })).toEqual({ invisibleCharacters: true })
  })

  it('recusa valor que não é booleano', () => {
    expect(schema.safeParse({ spellcheck: 'sim' }).success).toBe(false)
  })
})

describe('validação de file:save', () => {
  const schema = ipcContracts[IpcChannel.FileSave].request

  it('aceita uma gravação bem formada', () => {
    expect(schema.safeParse({ path: '/home/ana/ata.txt', content: 'texto' }).success).toBe(true)
  })

  it.each([
    ['sem caminho', { content: 'texto' }],
    ['caminho vazio', { path: '', content: 'texto' }],
    ['sem conteúdo', { path: '/a/b.txt' }],
    ['caminho não textual', { path: 42, content: 'texto' }],
    ['nulo', null],
  ])('recusa %s', (_label, payload) => {
    expect(schema.safeParse(payload).success).toBe(false)
  })

  it('recusa conteúdo acima do teto de memória', () => {
    const oversized = { path: '/a/b.txt', content: 'x'.repeat(MAX_TEXT_LENGTH + 1) }
    expect(schema.safeParse(oversized).success).toBe(false)
  })

  it('descarta campos não previstos no contrato', () => {
    const parsed = schema.parse({ path: '/a/b.txt', content: 'oi', extra: 'ignorar' })
    expect(parsed).toEqual({ path: '/a/b.txt', content: 'oi' })
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
  // Quem executa este schema na resposta de verdade é `src/main/ipc/registry.ts`,
  // e `registry.test.ts` prova que executa: um schema que ninguém roda é tipo em
  // tempo de compilação, não proteção.
  const schema = ipcContracts[IpcChannel.FontsList].response

  it('aceita a lista vazia', () => {
    // Sistema sem `fontconfig` devolve nada, e isso não é erro: a barra segue
    // com as fontes que o instalador leva.
    expect(schema.safeParse({ families: [] }).success).toBe(true)
  })

  it('recusa nome vazio e lista absurda', () => {
    // O que chega aqui é saída de programa do sistema. Ela é dado, não verdade:
    // um nome vazio viraria opção invisível no seletor, e uma lista de cem mil
    // entradas travaria a barra ao abrir.
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
