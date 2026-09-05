/**
 * Testes do cliente contra processos de verdade.
 *
 * Os cenários de falha usam sidecars de mentira escritos em Node: é a única
 * forma de provocar de propósito o que um sidecar real faz por acidente —
 * morrer no meio, emudecer, cuspir lixo. O sidecar .NET verdadeiro é exercitado
 * em `sidecar-real.test.ts`.
 */

import { mkdtemp, writeFile, chmod } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, describe, expect, it } from 'vitest'
import { ErrorCode, type AppError } from '@shared/errors.js'
import { SidecarClient } from './client.js'
import { SidecarMethod, encodeFrame } from './protocol.js'

/**
 * O sidecar de mentira é um script com shebang, e só POSIX o executa direto.
 *
 * O Windows não tem shebang: o `spawn` de um `.sh` devolve EFTYPE — "isto não é
 * uma imagem executável". Trocar por um `.cmd` não resolveria, porque o Node
 * recusa `.bat` e `.cmd` sem `shell: true`, e o cliente spawna com
 * `shell: false` de propósito; e alargar o resolvedor para aceitar argumentos
 * afrouxaria justamente o contrato que o wrapper existe para preservar — um
 * caminho só, sem argumentos, como o binário .NET de verdade.
 *
 * O que estes testes provam — decodificar quadros, expirar, subir de novo
 * depois de uma queda — é lógica sem plataforma nenhuma, e roda no Linux. O que
 * é de plataforma continua coberto no Windows: `sidecar-real.test.ts` conversa
 * com o `.exe` publicado, e o job do instalador exercita o aplicativo
 * empacotado.
 *
 * O portão vale só para quem precisa do processo **vivo**. Quem não precisa —
 * o executável que não existe, e os que encerram o cliente antes de pedir
 * qualquer coisa — continua rodando nos dois sistemas.
 */
const sidecarDeMentiraSobe = process.platform !== 'win32'

const clients: SidecarClient[] = []

afterEach(() => {
  for (const client of clients.splice(0)) client.dispose()
})

/** Escreve um sidecar de mentira e devolve um cliente já apontado para ele. */
async function fakeSidecar(source: string): Promise<SidecarClient> {
  const directory = await mkdtemp(join(tmpdir(), 'librevia-sidecar-'))
  const script = join(directory, 'fake.mjs')
  await writeFile(script, source, 'utf8')
  await chmod(script, 0o755)

  // O cliente executa um caminho só, sem argumentos — como fará com o binário
  // .NET. O wrapper existe para caber nesse contrato sem afrouxá-lo no teste.
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
    // Sem correlação por id, a resposta de um pedido chegaria para outro — e o
    // usuário veria o documento errado sem nenhum erro aparecer.
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
    // É o pior caso para o usuário: sem timeout, a janela congela e a única
    // saída é matar o aplicativo — perdendo o que não foi salvo.
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

    // O processo novo começa do zero, então `primeiro` volta a ser true e ele
    // morre de novo; o que importa é que houve uma segunda tentativa real.
    const segundo = await codeOf(client.request(SidecarMethod.Echo, {}))
    expect(segundo).toBe(ErrorCode.SidecarFailed)
  })

  it.runIf(sidecarDeMentiraSobe)('derruba um sidecar mudo em vez de deixá-lo acumulando pedidos', async () => {
    const client = await fakeSidecar(`${RESPONDER}
      function __handle() { /* nunca responde */ }
    `)

    await expect(client.request(SidecarMethod.Echo, {}, undefined, 200)).rejects.toThrow()
    // Se o processo tivesse sobrevivido ao timeout, o segundo pedido cairia no
    // mesmo laço travado. Ele precisa começar limpo.
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
    // Este é o contrato que o C# lê. Se mudar aqui e não lá, tudo quebra.
    const frame = encodeFrame({ a: 1 }, new Uint8Array([1, 2, 3]))
    const view = new DataView(frame.buffer, frame.byteOffset)

    expect(view.getUint32(0, false)).toBe(JSON.stringify({ a: 1 }).length)
    expect(view.getUint32(4, false)).toBe(3)
  })
})
