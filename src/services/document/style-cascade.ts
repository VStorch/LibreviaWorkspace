/**
 * A cascata dos estilos: o que um estilo **vale**, e não o que ele declara.
 *
 * A mesma ordem de `StyleResolver.cs` — padrões do documento, depois a cadeia de
 * `basedOn` do ancestral mais distante ao mais próximo —, porque é contra ele que
 * a tela vai ser comparada no dia em que o leitor parar de achatar: dois
 * resolvedores que discordam fazem o documento mudar de aparência ao reabrir.
 *
 * O que o sidecar ainda não resolve fica de fora daqui também, de propósito:
 * estilo de caractere (`w:rStyle`), o "inverte o herdado" das propriedades
 * liga/desliga, estilo de tabela e fontes de tema. Paridade antes de melhoria.
 */

import type { StyleCharacterFormat, StyleDefinition, StyleParagraphFormat, StyleSheet } from './styles.js'

/** Um estilo com a herança aplicada. Campo ausente é "ninguém na cadeia disse". */
export interface ResolvedStyle {
  readonly paragraph: StyleParagraphFormat
  readonly character: StyleCharacterFormat
}

/** O limite de `StyleResolver.MaxChainDepth`: cadeia mais longa é arquivo quebrado. */
const MAX_CHAIN_DEPTH = 16

const cache = new WeakMap<StyleSheet, Map<string, ResolvedStyle>>()

/**
 * O que o estilo vale — `null` é o estilo padrão de parágrafo.
 *
 * Um id que o documento não define resolve só os padrões, como no Word: o
 * parágrafo que aponta para o nada não herda de ninguém.
 */
export function resolveStyle(sheet: StyleSheet, styleId: string | null): ResolvedStyle {
  let resolved = cache.get(sheet)
  if (resolved === undefined) {
    resolved = new Map()
    cache.set(sheet, resolved)
  }

  const key = styleId ?? ''
  const cached = resolved.get(key)
  if (cached !== undefined) return cached

  let paragraph: StyleParagraphFormat = { ...sheet.defaults.paragraph }
  let character: StyleCharacterFormat = { ...sheet.defaults.character }
  for (const style of chainOf(sheet, styleId ?? sheet.defaults.paragraphStyleId)) {
    paragraph = overlay(paragraph, style.paragraph)
    character = overlay(character, style.character)
  }

  const result = { paragraph, character }
  resolved.set(key, result)
  return result
}

function chainOf(sheet: StyleSheet, styleId: string | null): readonly StyleDefinition[] {
  const chain: StyleDefinition[] = []
  const seen = new Set<string>()
  let current = styleId ?? undefined

  while (current !== undefined && !seen.has(current) && chain.length < MAX_CHAIN_DEPTH) {
    const style = sheet.styles[current]
    if (style === undefined) break
    seen.add(current)
    chain.push(style)
    current = style.basedOn
  }

  return chain.reverse()
}

/**
 * Campo a campo, e só o que foi dito: o silêncio de um estilo deixa passar o
 * herdado. É o equivalente do `AttributeByAttribute` do C# — aqui antes, depois
 * e entrelinha já são campos separados, então o `w:spacing` que só redeclara o
 * espaço não apaga a entrelinha.
 */
function overlay<T extends object>(base: T, top: T | undefined): T {
  if (top === undefined) return base
  const merged: Record<string, unknown> = { ...(base as Record<string, unknown>) }
  for (const [key, value] of Object.entries(top)) {
    if (value !== undefined) merged[key] = value
  }
  return merged as T
}
