import { ALLOWED_EXTERNAL_PROTOCOLS } from '@shared/constants.js'

/**
 * Normaliza e valida o endereço de um link.
 *
 * Um documento pode vir de qualquer lugar, e `javascript:` num link é execução
 * de código disfarçada de texto. A allowlist de esquemas é a mesma que o
 * processo main aplica antes de abrir qualquer coisa no navegador do sistema —
 * esta função só evita que um endereço inválido chegue a ser gravado.
 *
 * Devolve `null` quando o endereço não pode ser aceito.
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
