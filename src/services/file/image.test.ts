import { describe, expect, it } from 'vitest'
import { MAX_IMAGE_BYTES, detectImageMimeType, isImageWithinSizeLimit } from './image.js'

const bytesOf = (...values: number[]): Uint8Array => Uint8Array.from(values)

describe('detectImageMimeType', () => {
  it('reconhece PNG', () => {
    expect(detectImageMimeType(bytesOf(0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00))).toBe(
      'image/png',
    )
  })

  it('reconhece JPEG', () => {
    expect(detectImageMimeType(bytesOf(0xff, 0xd8, 0xff, 0xe0, 0x00))).toBe('image/jpeg')
  })

  it('reconhece GIF', () => {
    expect(detectImageMimeType(bytesOf(0x47, 0x49, 0x46, 0x38, 0x39, 0x61))).toBe('image/gif')
  })

  it('reconhece WebP', () => {
    const webp = bytesOf(0x52, 0x49, 0x46, 0x46, 0, 0, 0, 0, 0x57, 0x45, 0x42, 0x50)
    expect(detectImageMimeType(webp)).toBe('image/webp')
  })

  it('não confunde outro contêiner RIFF com WebP', () => {
    // WAV também começa com "RIFF"; só os bytes 8..11 distinguem.
    const wav = bytesOf(0x52, 0x49, 0x46, 0x46, 0, 0, 0, 0, 0x57, 0x41, 0x56, 0x45)
    expect(detectImageMimeType(wav)).toBeNull()
  })

  it('recusa SVG mesmo sendo imagem', () => {
    // SVG é XML e pode carregar script: embuti-lo seria XSS dentro do
    // documento do próprio usuário.
    const svg = new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg"><script/></svg>')
    expect(detectImageMimeType(svg)).toBeNull()
  })

  it('recusa executável renomeado como imagem', () => {
    expect(detectImageMimeType(bytesOf(0x7f, 0x45, 0x4c, 0x46))).toBeNull()
    expect(detectImageMimeType(bytesOf(0x4d, 0x5a, 0x90, 0x00))).toBeNull()
  })

  it('recusa conteúdo vazio ou curto demais para ter assinatura', () => {
    expect(detectImageMimeType(bytesOf())).toBeNull()
    expect(detectImageMimeType(bytesOf(0x89, 0x50))).toBeNull()
  })
})

describe('isImageWithinSizeLimit', () => {
  it.each([
    [1, true],
    [MAX_IMAGE_BYTES, true],
    [MAX_IMAGE_BYTES + 1, false],
    [0, false],
  ])('para %i bytes devolve %s', (size, expected) => {
    expect(isImageWithinSizeLimit(size)).toBe(expected)
  })
})
