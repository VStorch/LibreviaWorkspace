/**
 * Reconhecimento de imagens por assinatura de bytes.
 *
 * A extensão do arquivo não é evidência: um `.png` pode conter qualquer coisa.
 * Como a imagem vai virar um data URI dentro do documento, o que decide é o
 * conteúdo real.
 *
 * SVG é recusado de propósito, mesmo sendo uma imagem legítima: é um documento
 * XML que pode carregar script, e embuti-lo num editor seria abrir XSS dentro
 * do próprio arquivo do usuário (ver docs/00-plano-tecnico.md §6.7).
 */

export const ALLOWED_IMAGE_MIME_TYPES = ['image/png', 'image/jpeg', 'image/gif', 'image/webp'] as const
export type AllowedImageMimeType = (typeof ALLOWED_IMAGE_MIME_TYPES)[number]

export const MAX_IMAGE_BYTES = 10 * 1024 * 1024

function startsWith(bytes: Uint8Array, signature: readonly number[], offset = 0): boolean {
  if (bytes.length < offset + signature.length) return false
  return signature.every((byte, index) => bytes[offset + index] === byte)
}

/** Tipo real da imagem, ou `null` se não for um formato aceito. */
export function detectImageMimeType(bytes: Uint8Array): AllowedImageMimeType | null {
  if (startsWith(bytes, [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])) return 'image/png'
  if (startsWith(bytes, [0xff, 0xd8, 0xff])) return 'image/jpeg'
  if (startsWith(bytes, [0x47, 0x49, 0x46, 0x38])) return 'image/gif'

  // WebP: "RIFF" .... "WEBP"
  if (startsWith(bytes, [0x52, 0x49, 0x46, 0x46]) && startsWith(bytes, [0x57, 0x45, 0x42, 0x50], 8)) {
    return 'image/webp'
  }

  return null
}

export function isImageWithinSizeLimit(byteLength: number): boolean {
  return byteLength > 0 && byteLength <= MAX_IMAGE_BYTES
}
