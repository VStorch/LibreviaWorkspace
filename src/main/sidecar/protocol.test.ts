import { describe, expect, it } from 'vitest'
import { AppError, ErrorCode } from '@shared/errors.js'
import {
  EMPTY_BINARY,
  FRAME_HEADER_BYTES,
  FrameReader,
  MAX_BINARY_BYTES,
  encodeFrame,
  parseResponse,
} from './protocol.js'

const readAll = (chunks: Uint8Array[]) => {
  const reader = new FrameReader()
  return chunks.flatMap((chunk) => reader.push(chunk))
}

describe('ida e volta do quadro', () => {
  it('recupera JSON e binário exatamente como entraram', () => {
    const binary = new Uint8Array([0, 255, 13, 10, 0, 127])
    const [frame] = readAll([encodeFrame({ id: 1, method: 'health' }, binary)])

    expect(frame?.json).toEqual({ id: 1, method: 'health' })
    expect(frame?.binary).toEqual(binary)
  })

  it('aceita quadro sem binário', () => {
    const [frame] = readAll([encodeFrame({ ok: true })])

    expect(frame?.json).toEqual({ ok: true })
    expect(frame?.binary).toEqual(EMPTY_BINARY)
  })

  it('preserva bytes que pareceriam fim de linha ou fim de texto', () => {
    // Um protocolo delimitado por \n se despedaçaria aqui. O nosso é por
    // tamanho justamente para carregar DOCX, que é ZIP e contém de tudo.
    const binary = new Uint8Array([0x0a, 0x0d, 0x1a, 0x00, 0x50, 0x4b, 0x03, 0x04])
    const [frame] = readAll([encodeFrame({}, binary)])

    expect(frame?.binary).toEqual(binary)
  })

  it('preserva texto não-ASCII', () => {
    const [frame] = readAll([encodeFrame({ message: 'Configuração — evidências' })])

    expect(frame?.json).toEqual({ message: 'Configuração — evidências' })
  })
})

describe('fronteiras de mensagem no pipe', () => {
  // Um pipe não preserva fronteiras. Estes três casos são o bug clássico
  // desta integração, e ele só aparece com documento grande.

  it('remonta um quadro partido byte a byte', () => {
    const complete = encodeFrame({ id: 7 }, new Uint8Array([1, 2, 3, 4, 5]))
    const frames = readAll([...complete].map((byte) => new Uint8Array([byte])))

    expect(frames).toHaveLength(1)
    expect(frames[0]?.json).toEqual({ id: 7 })
    expect(frames[0]?.binary).toEqual(new Uint8Array([1, 2, 3, 4, 5]))
  })

  it('separa vários quadros que chegaram num pedaço só', () => {
    const a = encodeFrame({ id: 1 })
    const b = encodeFrame({ id: 2 }, new Uint8Array([9]))
    const c = encodeFrame({ id: 3 })

    const glued = new Uint8Array(a.length + b.length + c.length)
    glued.set(a, 0)
    glued.set(b, a.length)
    glued.set(c, a.length + b.length)

    expect(readAll([glued]).map((f) => f.json)).toEqual([{ id: 1 }, { id: 2 }, { id: 3 }])
  })

  it('não entrega nada enquanto o quadro está incompleto', () => {
    const reader = new FrameReader()
    const complete = encodeFrame({ id: 1 }, new Uint8Array([1, 2, 3]))

    expect(reader.push(complete.subarray(0, FRAME_HEADER_BYTES))).toEqual([])
    expect(reader.push(complete.subarray(FRAME_HEADER_BYTES, complete.length - 1))).toEqual([])
    expect(reader.push(complete.subarray(complete.length - 1))).toHaveLength(1)
  })

  it('sobrevive a um binário de 1 MB partido em pedaços de 64 KB', () => {
    const binary = new Uint8Array(1024 * 1024).map((_, i) => i % 256)
    const complete = encodeFrame({ id: 1 }, binary)

    const chunks: Uint8Array[] = []
    for (let at = 0; at < complete.length; at += 65536) {
      chunks.push(complete.subarray(at, at + 65536))
    }

    const frames = readAll(chunks)
    expect(frames).toHaveLength(1)
    expect(frames[0]?.binary).toEqual(binary)
  })
})

describe('defesa contra sidecar corrompido', () => {
  it('recusa quadro que anuncia binário absurdo, sem tentar alocar', () => {
    // Sem este teto, um cabeçalho mentiroso nos faria reservar gigabytes.
    const header = new Uint8Array(FRAME_HEADER_BYTES)
    new DataView(header.buffer).setUint32(4, MAX_BINARY_BYTES + 1, false)

    expect(() => new FrameReader().push(header)).toThrow(AppError)
    expect(() => new FrameReader().push(header)).toThrow(/inesperada/i)
  })

  it('recusa JSON malformado no quadro', () => {
    const garbage = new TextEncoder().encode('{ isto não é json')
    const frame = new Uint8Array(FRAME_HEADER_BYTES + garbage.length)
    new DataView(frame.buffer).setUint32(0, garbage.length, false)
    frame.set(garbage, FRAME_HEADER_BYTES)

    expect(() => new FrameReader().push(frame)).toThrow(/inesperada/i)
  })
})

describe('parseResponse', () => {
  it('aceita sucesso', () => {
    const response = parseResponse({ id: 1, ok: true, result: { version: '1.0' } })

    expect(response.ok).toBe(true)
    expect(response).toMatchObject({ id: 1 })
  })

  it('aceita falha com erro do sidecar', () => {
    const response = parseResponse({
      id: 2,
      ok: false,
      error: { code: 'DOCX_INVALID', message: 'O arquivo não é um documento válido.' },
    })

    expect(response.ok).toBe(false)
  })

  it.each([
    ['sem id', { ok: true, result: 1 }],
    ['sem ok', { id: 1, result: 1 }],
    ['erro sem mensagem', { id: 1, ok: false, error: { code: 'X' } }],
    ['id negativo', { id: -1, ok: true, result: 1 }],
    ['não é objeto', 'ok'],
  ])('recusa resposta fora do contrato: %s', (_label, payload) => {
    // O sidecar é outro processo: pode estar numa versão antiga ou corrompido.
    // Confiar na forma do que ele devolve seria o erro de confiar no renderer.
    expect(() => parseResponse(payload)).toThrow(AppError)
    try {
      parseResponse(payload)
    } catch (error) {
      expect((error as AppError).code).toBe(ErrorCode.SidecarFailed)
    }
  })

  it('não vaza detalhe técnico na mensagem ao usuário', () => {
    try {
      parseResponse({ nada: 'disso' })
    } catch (error) {
      expect((error as AppError).message).not.toMatch(/zod|schema|undefined/i)
    }
  })
})
