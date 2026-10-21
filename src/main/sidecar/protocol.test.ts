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

describe('frame round trip', () => {
  it('recovers JSON and binary exactly as they went in', () => {
    const binary = new Uint8Array([0, 255, 13, 10, 0, 127])
    const [frame] = readAll([encodeFrame({ id: 1, method: 'health' }, binary)])

    expect(frame?.json).toEqual({ id: 1, method: 'health' })
    expect(frame?.binary).toEqual(binary)
  })

  it('accepts a frame without binary', () => {
    const [frame] = readAll([encodeFrame({ ok: true })])

    expect(frame?.json).toEqual({ ok: true })
    expect(frame?.binary).toEqual(EMPTY_BINARY)
  })

  it('keeps bytes that would look like end of line or end of text', () => {
    // A newline-delimited protocol would fall apart here. Ours is length-prefixed precisely to
    // carry DOCX, which is a ZIP and contains anything.
    const binary = new Uint8Array([0x0a, 0x0d, 0x1a, 0x00, 0x50, 0x4b, 0x03, 0x04])
    const [frame] = readAll([encodeFrame({}, binary)])

    expect(frame?.binary).toEqual(binary)
  })

  it('keeps non-ASCII text', () => {
    const [frame] = readAll([encodeFrame({ message: 'Configuração — evidências' })])

    expect(frame?.json).toEqual({ message: 'Configuração — evidências' })
  })
})

describe('message boundaries in the pipe', () => {
  // A pipe does not preserve boundaries. These three cases are this integration's classic bug, and
  // it only shows with large documents.

  it('reassembles a frame split byte by byte', () => {
    const complete = encodeFrame({ id: 7 }, new Uint8Array([1, 2, 3, 4, 5]))
    const frames = readAll([...complete].map((byte) => new Uint8Array([byte])))

    expect(frames).toHaveLength(1)
    expect(frames[0]?.json).toEqual({ id: 7 })
    expect(frames[0]?.binary).toEqual(new Uint8Array([1, 2, 3, 4, 5]))
  })

  it('splits several frames that arrived in one chunk', () => {
    const a = encodeFrame({ id: 1 })
    const b = encodeFrame({ id: 2 }, new Uint8Array([9]))
    const c = encodeFrame({ id: 3 })

    const glued = new Uint8Array(a.length + b.length + c.length)
    glued.set(a, 0)
    glued.set(b, a.length)
    glued.set(c, a.length + b.length)

    expect(readAll([glued]).map((f) => f.json)).toEqual([{ id: 1 }, { id: 2 }, { id: 3 }])
  })

  it('delivers nothing while the frame is incomplete', () => {
    const reader = new FrameReader()
    const complete = encodeFrame({ id: 1 }, new Uint8Array([1, 2, 3]))

    expect(reader.push(complete.subarray(0, FRAME_HEADER_BYTES))).toEqual([])
    expect(reader.push(complete.subarray(FRAME_HEADER_BYTES, complete.length - 1))).toEqual([])
    expect(reader.push(complete.subarray(complete.length - 1))).toHaveLength(1)
  })

  it('survives a 1 MB binary split into 64 KB chunks', () => {
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

describe('defense against a corrupt sidecar', () => {
  it('refuses a frame announcing an absurd binary, without trying to allocate', () => {
    // Without this cap, a lying header would make us reserve gigabytes.
    const header = new Uint8Array(FRAME_HEADER_BYTES)
    new DataView(header.buffer).setUint32(4, MAX_BINARY_BYTES + 1, false)

    expect(() => new FrameReader().push(header)).toThrow(AppError)
    expect(() => new FrameReader().push(header)).toThrow(/inesperada/i)
  })

  it('refuses malformed JSON in the frame', () => {
    const garbage = new TextEncoder().encode('{ this is not json')
    const frame = new Uint8Array(FRAME_HEADER_BYTES + garbage.length)
    new DataView(frame.buffer).setUint32(0, garbage.length, false)
    frame.set(garbage, FRAME_HEADER_BYTES)

    expect(() => new FrameReader().push(frame)).toThrow(/inesperada/i)
  })
})

describe('parseResponse', () => {
  it('accepts success', () => {
    const response = parseResponse({ id: 1, ok: true, result: { version: '1.0' } })

    expect(response.ok).toBe(true)
    expect(response).toMatchObject({ id: 1 })
  })

  it('accepts a failure with a sidecar error', () => {
    const response = parseResponse({
      id: 2,
      ok: false,
      error: { code: 'DOCX_INVALID', message: 'O arquivo não é um documento válido.' },
    })

    expect(response.ok).toBe(false)
  })

  it.each([
    ['without id', { ok: true, result: 1 }],
    ['without ok', { id: 1, result: 1 }],
    ['error without message', { id: 1, ok: false, error: { code: 'X' } }],
    ['negative id', { id: -1, ok: true, result: 1 }],
    ['not an object', 'ok'],
  ])('refuses a response outside the contract: %s', (_label, payload) => {
    // The sidecar is another process: it may be an old version or corrupt. Trusting the shape of
    // what it returns would be the mistake of trusting the renderer.
    expect(() => parseResponse(payload)).toThrow(AppError)
    try {
      parseResponse(payload)
    } catch (error) {
      expect((error as AppError).code).toBe(ErrorCode.SidecarFailed)
    }
  })

  it('does not leak technical detail in the user message', () => {
    try {
      parseResponse({ nada: 'disso' })
    } catch (error) {
      expect((error as AppError).message).not.toMatch(/zod|schema|undefined/i)
    }
  })
})
