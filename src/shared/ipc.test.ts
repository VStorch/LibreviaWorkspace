import { describe, expect, it } from 'vitest'
import { ALL_IPC_CHANNELS, IpcChannel } from './ipc-channels.js'
import { appPingRequestSchema, ipcContracts } from './ipc.js'

describe('contratos de IPC', () => {
  it('define um contrato para cada canal declarado', () => {
    // Um canal sem schema passaria payload não validada ao handler.
    for (const channel of ALL_IPC_CHANNELS) {
      expect(ipcContracts[channel]).toBeDefined()
    }
  })

  it('não expõe contrato para canal não declarado', () => {
    expect(Object.keys(ipcContracts).sort()).toEqual([...ALL_IPC_CHANNELS].sort())
  })
})

describe('validação de payload', () => {
  it('aceita uma requisição bem formada', () => {
    expect(appPingRequestSchema.safeParse({ message: 'olá' }).success).toBe(true)
  })

  it.each([
    ['objeto vazio', {}],
    ['mensagem vazia', { message: '' }],
    ['tipo errado', { message: 42 }],
    ['nulo', null],
    ['string solta', 'olá'],
    ['acima do limite', { message: 'x'.repeat(201) }],
  ])('recusa %s', (_label, payload) => {
    expect(appPingRequestSchema.safeParse(payload).success).toBe(false)
  })

  it('descarta campos não previstos no contrato', () => {
    const parsed = appPingRequestSchema.parse({ message: 'olá', extra: 'ignorar' })
    expect(parsed).toEqual({ message: 'olá' })
  })

  it('nomeia o canal de fumaça de forma estável', () => {
    expect(IpcChannel.AppPing).toBe('app:ping')
  })
})
