import { describe, expect, it, vi } from 'vitest'
import { IpcChannel } from '@shared/ipc-channels.js'
import { ErrorCode } from '@shared/errors.js'

/**
 * O que o registro promete, cobrado nas duas direções.
 *
 * O pedido vem do renderer e é tratado como não confiável — isso o registro já
 * fazia. A **resposta** vem de casa, e por muito tempo ninguém a conferia: o
 * `response` de cada canal era tipo em tempo de compilação e mais nada. Um
 * handler que devolvesse a forma errada — nome de fonte vazio saído do `fc-list`,
 * campo esquecido depois de um refatoramento — chegava inteiro à interface, e o
 * defeito aparecia longe da causa.
 *
 * O `ipcMain` é falsificado porque não há Electron aqui: o que interessa é a
 * função que o registro entrega a ele.
 */
const { handlers } = vi.hoisted(() => ({
  handlers: new Map<string, (event: unknown, payload: unknown) => Promise<unknown>>(),
}))

vi.mock('electron', () => ({
  ipcMain: {
    handle: (channel: string, handler: (event: unknown, payload: unknown) => Promise<unknown>) => {
      handlers.set(channel, handler)
    },
  },
}))

const { handle } = await import('./registry.js')

/** Registra o canal com um handler de mentira e chama o que o `ipcMain` chamaria. */
async function invoke(channel: IpcChannel, reply: unknown, payload: unknown = {}): Promise<unknown> {
  // O tipo do handler é o do contrato; aqui o teste devolve de propósito o que o
  // contrato não prevê, que é o caso a cobrir.
  handle(channel as never, (() => reply) as never)
  return handlers.get(channel)!({}, payload)
}

describe('registro de IPC', () => {
  it('devolve a resposta válida já normalizada', () => {
    expect(invoke(IpcChannel.FontsList, { families: ['Arial'] })).resolves.toEqual({
      ok: true,
      data: { families: ['Arial'] },
    })
  })

  it('recusa a resposta que o contrato do canal não prevê', async () => {
    // Nome vazio é o que uma saída estragada do `fc-list` produz, e viraria opção
    // invisível no seletor. O renderer recebe erro, e o log do main diz por quê.
    const result = (await invoke(IpcChannel.FontsList, { families: [''] })) as {
      ok: boolean
      error?: { code: string }
    }

    expect(result.ok).toBe(false)
    expect(result.error?.code).toBe(ErrorCode.Internal)
  })

  it('recusa o pedido que o contrato do canal não prevê', async () => {
    const result = (await invoke(
      IpcChannel.SpellAddWord,
      { added: true },
      { word: '', scope: 'session' },
    )) as {
      ok: boolean
      error?: { code: string }
    }

    expect(result.ok).toBe(false)
    expect(result.error?.code).toBe(ErrorCode.InvalidRequest)
  })
})
