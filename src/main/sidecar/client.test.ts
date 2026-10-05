/**
 * The client against real processes. Failures use fake Node sidecars that die, go silent and spit
 * garbage on purpose; the real .NET one is in `sidecar-real.test.ts`.
 */

import { mkdtemp, writeFile, chmod } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import { ErrorCode, type AppError } from '@shared/errors.js'
import { SidecarClient } from './client.js'
import { SidecarMethod, encodeFrame } from './protocol.js'

/**
 * The fake sidecar is a script with a shebang, which only POSIX runs: on Windows `spawn` gives
 * EFTYPE, Node refuses `.cmd` without `shell: true`, and accepting arguments would loosen the
 * contract. The logic under test has no platform; Windows is covered by `sidecar-real.test.ts` and
 * the installer job. Only the tests that need a live process depend on this.
 */
const sidecarDeMentiraSobe = process.platform !== 'win32'

const clients: SidecarClient[] = []

afterEach(() => {
  for (const client of clients.splice(0)) client.dispose()
})

async function fakeSidecar(source: string): Promise<SidecarClient> {
  const directory = await mkdtemp(join(tmpdir(), 'librevia-sidecar-'))
  const script = join(directory, 'fake.mjs')
  await writeFile(script, source, 'utf8')
  await chmod(script, 0o755)

  // The client runs a single path without arguments, like the .NET binary.
  const wrapper = join(directory, 'run.sh')
  await writeFile(wrapper, `#!/bin/sh\nexec "${process.execPath}" "${script}"\n`, 'utf8')
  await chmod(wrapper, 0o755)

  const client = new SidecarClient(() => Promise.resolve(wrapper))
  clients.push(client)
  return client
}

const codeOf = async (promise: Promise<unknown>): Promise<string> => {
  try {
    await promise
    throw new Error('esperava falha')
  } catch (error) {
    return (error as AppError).code
  }
}

const RESPONDER = `
import { Buffer } from 'node:buffer'
let buffer = Buffer.alloc(0)
process.stdin.on('data', (chunk) => {
  buffer = Buffer.concat([buffer, chunk])
  for (;;) {
    if (buffer.length < 8) return
    const jsonLength = buffer.readUInt32BE(0)
    const binaryLength = buffer.readUInt32BE(4)
    if (buffer.length < 8 + jsonLength + binaryLength) return
    const request = JSON.parse(buffer.subarray(8, 8 + jsonLength).toString('utf8'))
    const binary = buffer.subarray(8 + jsonLength, 8 + jsonLength + binaryLength)
    buffer = buffer.subarray(8 + jsonLength + binaryLength)
    __handle(request, binary)
  }
})
function reply(json, binary = Buffer.alloc(0)) {
  const body = Buffer.from(JSON.stringify(json), 'utf8')
  const header = Buffer.alloc(8)
  header.writeUInt32BE(body.length, 0)
  header.writeUInt32BE(binary.length, 4)
  process.stdout.write(Buffer.concat([header, body, binary]))
}
`

describe.runIf(sidecarDeMentiraSobe)('conversa normal', () => {
  it('responde a um pedido e devolve o binário', async () => {
    const client = await fakeSidecar(`${RESPONDER}
      function __handle(request, binary) {
        reply({ id: request.id, ok: true, result: { eco: request.params.valor } }, binary)
      }
    `)

    const reply = await client.request(SidecarMethod.Echo, { valor: 42 }, new Uint8Array([7, 8, 9]))

    expect(reply.result).toEqual({ eco: 42 })
    expect(reply.binary).toEqual(new Uint8Array([7, 8, 9]))
  })

  it('mantém pedidos simultâneos separados, mesmo respondidos fora de ordem', async () => {
    // Without correlation by id, one request's response would go to another, without error.
    const client = await fakeSidecar(`${RESPONDER}
      const pending = []
      function __handle(request) {
        pending.push(request)
        if (pending.length < 3) return
        for (const r of pending.reverse()) reply({ id: r.id, ok: true, result: r.params.n })
      }
    `)

    const results = await Promise.all([
      client.request(SidecarMethod.Echo, { n: 1 }),
      client.request(SidecarMethod.Echo, { n: 2 }),
      client.request(SidecarMethod.Echo, { n: 3 }),
    ])

    expect(results.map((r) => r.result)).toEqual([1, 2, 3])
  })

  it('propaga o erro do sidecar com a frase que ele mandou', async () => {
    const client = await fakeSidecar(`${RESPONDER}
      function __handle(request) {
        reply({ id: request.id, ok: false, error: { code: 'DOCX_INVALIDO', message: 'O arquivo não é um documento do Word.' } })
      }
    `)

    await expect(client.request(SidecarMethod.Echo, {})).rejects.toThrow(
      'O arquivo não é um documento do Word.',
    )
  })
})

describe('o sidecar morre — o documento não pode morrer junto', () => {
  it.runIf(sidecarDeMentiraSobe)('falha com erro compreensível quando o processo morre no meio', async () => {
    const client = await fakeSidecar(`${RESPONDER}
      function __handle() { process.exit(1) }
    `)

    const error = await codeOf(client.request(SidecarMethod.Echo, {}))

    expect(error).toBe(ErrorCode.SidecarFailed)
  })

  it.runIf(sidecarDeMentiraSobe)('não deixa o pedido pendurado para sempre quando o sidecar emudece', async () => {
    // Without a timeout, the window freezes and the only way out is killing the app.
    const client = await fakeSidecar(`${RESPONDER}
      function __handle() { /* nunca responde */ }
    `)

    const error = await codeOf(client.request(SidecarMethod.Echo, {}, undefined, 300))

    expect(error).toBe(ErrorCode.SidecarTimeout)
  })

  it.runIf(sidecarDeMentiraSobe)('sobe um processo novo depois de uma queda, em vez de ficar inutilizado', async () => {
    const client = await fakeSidecar(`${RESPONDER}
      let primeiro = true
      function __handle(request) {
        if (primeiro) { primeiro = false; process.exit(1) }
        reply({ id: request.id, ok: true, result: 'vivo' })
      }
    `)

    await expect(client.request(SidecarMethod.Echo, {})).rejects.toThrow()

    // The new process dies again; what matters is that there was a second attempt.
    const segundo = await codeOf(client.request(SidecarMethod.Echo, {}))
    expect(segundo).toBe(ErrorCode.SidecarFailed)
  })

  it.runIf(sidecarDeMentiraSobe)('derruba um sidecar mudo em vez de deixá-lo acumulando pedidos', async () => {
    const client = await fakeSidecar(`${RESPONDER}
      function __handle() { /* nunca responde */ }
    `)

    await expect(client.request(SidecarMethod.Echo, {}, undefined, 200)).rejects.toThrow()
    // The second request starts clean, not in the stuck loop.
    await expect(client.request(SidecarMethod.Echo, {}, undefined, 200)).rejects.toThrow()
  })

  it('recusa quando o executável não existe, sem derrubar o aplicativo', async () => {
    const client = new SidecarClient(() => Promise.resolve('/caminho/que/nao/existe'))
    clients.push(client)

    const error = await codeOf(client.request(SidecarMethod.Echo, {}))

    expect(error).toBe(ErrorCode.SidecarUnavailable)
  })
})

describe.runIf(sidecarDeMentiraSobe)('o sidecar responde besteira', () => {
  it('recusa resposta fora do contrato', async () => {
    const client = await fakeSidecar(`${RESPONDER}
      function __handle() { reply({ isto: 'não é uma resposta' }) }
    `)

    expect(await codeOf(client.request(SidecarMethod.Echo, {}))).toBe(ErrorCode.SidecarFailed)
  })

  it('recusa fluxo corrompido sem tentar adivinhar onde o próximo quadro começa', async () => {
    const client = await fakeSidecar(`${RESPONDER}
      function __handle() { process.stdout.write(Buffer.from('lixo que não é quadro nenhum')) }
    `)

    expect(await codeOf(client.request(SidecarMethod.Echo, {}, undefined, 1500))).toBe(
      ErrorCode.SidecarFailed,
    )
  })

  it('ignora resposta com id que ninguém pediu', async () => {
    const client = await fakeSidecar(`${RESPONDER}
      function __handle(request) {
        reply({ id: 9999, ok: true, result: 'fantasma' })
        reply({ id: request.id, ok: true, result: 'certo' })
      }
    `)

    const reply = await client.request(SidecarMethod.Echo, {})

    expect(reply.result).toBe('certo')
  })
})

describe('encerramento', () => {
  it.runIf(sidecarDeMentiraSobe)('dispose não deixa pedido pendurado', async () => {
    const client = await fakeSidecar(`${RESPONDER}
      function __handle() { /* nunca responde */ }
    `)

    const pending = client.request(SidecarMethod.Echo, {})
    client.dispose()

    expect(await codeOf(pending)).toBe(ErrorCode.SidecarUnavailable)
  })

  it('recusa pedidos depois de encerrado', async () => {
    const client = await fakeSidecar(RESPONDER + 'function __handle() {}')
    client.dispose()

    expect(await codeOf(client.request(SidecarMethod.Echo, {}))).toBe(ErrorCode.SidecarUnavailable)
  })

  it('dispose é idempotente', async () => {
    const client = await fakeSidecar(RESPONDER + 'function __handle() {}')

    expect(() => {
      client.dispose()
      client.dispose()
    }).not.toThrow()
  })
})

describe('encodeFrame no formato que o sidecar espera', () => {
  it('põe os tamanhos em big-endian nos primeiros 8 bytes', () => {
    // The contract the C# side reads.
    const frame = encodeFrame({ a: 1 }, new Uint8Array([1, 2, 3]))
    const view = new DataView(frame.buffer, frame.byteOffset)

    expect(view.getUint32(0, false)).toBe(JSON.stringify({ a: 1 }).length)
    expect(view.getUint32(4, false)).toBe(3)
  })
})
