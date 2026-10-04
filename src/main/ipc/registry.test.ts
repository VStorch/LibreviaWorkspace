import { describe, expect, it, vi } from 'vitest'
import { IpcChannel } from '@shared/ipc-channels.js'
import { ErrorCode } from '@shared/errors.js'

/**
 * O registro cobra nas duas direções: o pedido do renderer não é confiável, e a
 * resposta de casa também é conferida, para a forma errada não chegar à interface
 * longe da causa. O `ipcMain` é falsificado, porque não há Electron aqui.
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
