import { ALLOWED_EXTERNAL_PROTOCOLS } from '@shared/constants.js'

/**
 * `javascript:` num link é código disfarçado. O main aplica a mesma lista de
 * esquemas antes de abrir; aqui só se evita gravar endereço inválido.
 */
export function normalizeLinkUrl(input: string): string | null {
  const trimmed = input.trim()
  if (trimmed.length === 0) return null

  // Endereço digitado sem esquema é o caso comum: "empresa.com.br".
  const candidate = /^[a-z][a-z0-9+.-]*:/i.test(trimmed) ? trimmed : `https://${trimmed}`

  let parsed: URL
  try {
    parsed = new URL(candidate)
  } catch {
    return null
  }

  if (!(ALLOWED_EXTERNAL_PROTOCOLS as readonly string[]).includes(parsed.protocol)) return null
  if (parsed.protocol !== 'mailto:' && parsed.hostname === '') return null

  return parsed.toString()
}
