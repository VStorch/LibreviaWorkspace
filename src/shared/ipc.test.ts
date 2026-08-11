import { describe, expect, it } from 'vitest'
import { INVOCABLE_IPC_CHANNELS, IpcChannel } from './ipc-channels.js'
import { MAX_TEXT_LENGTH, ipcContracts } from './ipc.js'

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

describe('validação de dialog:confirm-discard', () => {
  const response = ipcContracts[IpcChannel.DialogConfirmDiscard].response

  it.each(['save', 'discard', 'cancel'])('aceita a resposta %s', (choice) => {
    expect(response.safeParse({ choice }).success).toBe(true)
  })

  it('recusa uma resposta fora das três previstas', () => {
    expect(response.safeParse({ choice: 'talvez' }).success).toBe(false)
  })
})
