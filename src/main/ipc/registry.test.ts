import { describe, expect, it, vi } from 'vitest'
import { IpcChannel } from '@shared/ipc-channels.js'
import { ErrorCode } from '@shared/errors.js'

/**
 * The registry checks both directions: the renderer request is untrusted, and our own response is
 * checked too, so a wrong shape does not reach the UI far from its cause. `ipcMain` is faked, since
 * there is no Electron here.
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

async function invoke(channel: IpcChannel, reply: unknown, payload: unknown = {}): Promise<unknown> {
  // The handler type is the contract's; here the test returns on purpose what the contract does not
  // allow, which is the case to cover.
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
    // An empty name is what broken `fc-list` output produces, and would become an invisible option.
    // The renderer gets an error, and the main log says why.
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
