/**
 * A mesma ordem de `StyleResolver.cs`: padrões do documento, depois a cadeia de
 * `basedOn` do ancestral mais distante ao mais próximo. Dois resolvedores que
 * discordam mudam a aparência do documento ao reabrir. O que o sidecar não
 * resolve fica de fora daqui também: `w:rStyle`, o "inverte o herdado" das
 * propriedades liga/desliga, estilo de tabela e fontes de tema.
 */

import { cssLineHeightOf } from './line-metrics.js'
import {
  LEGACY_STYLES,
  StyleType,
  type LineSpacing,
  type StyleCharacterFormat,
  type StyleDefinition,
  type StyleParagraphFormat,
  type StyleSheet,
} from './styles.js'

/** Um estilo com a herança aplicada. Campo ausente é "ninguém na cadeia disse". */
export interface ResolvedStyle {
  readonly paragraph: StyleParagraphFormat
  readonly character: StyleCharacterFormat
}

/** O limite de `StyleResolver.MaxChainDepth`: cadeia mais longa é arquivo quebrado. */
const MAX_CHAIN_DEPTH = 16

const cache = new WeakMap<StyleSheet, Map<string, ResolvedStyle>>()

/** `null` é o estilo padrão. Um id que o documento não define resolve só os padrões, como no Word. */
export function resolveStyle(sheet: StyleSheet, styleId: string | null): ResolvedStyle {
  let resolved = cache.get(sheet)
  if (resolved === undefined) {
    resolved = new Map()
    cache.set(sheet, resolved)
  }

  // O nulo não pode dividir chave com id nenhum — nem com o vazio, que é o jeito
  // de pedir só os padrões (ver `style-css.ts`).
  const key = styleId ?? '\u0000'
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
 * Campo a campo: o silêncio de um estilo deixa passar o herdado. Como antes,
 * depois e entrelinha são campos separados, o `w:spacing` que só redeclara o
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

/** Sem o estilo no documento, o que o escritor vai acrescentar ao gravar (`BuiltinStyles.cs`). */
export function headingStyleOf(sheet: StyleSheet, level: number): ResolvedStyle {
  const name = `heading ${level}`
  const own = Object.values(sheet.styles).find(
    (style: StyleDefinition) => style.type === StyleType.Paragraph && style.name.toLowerCase() === name,
  )
  if (own !== undefined) return resolveStyle(sheet, own.id)

  const fallback = LEGACY_STYLES.styles[`Heading${level}`]
  const base = resolveStyle(sheet, null)
  return {
    paragraph: { ...base.paragraph, ...fallback?.paragraph },
    character: { ...base.character, ...fallback?.character },
  }
}

export interface StyledBlock {
  readonly type: { readonly name: string } | string
  readonly attrs?: Readonly<Record<string, unknown>> | null | undefined
}

/**
 * O estilo do bloco com a formatação direta por cima, nas unidades do nó. Quem
 * decide olhando atributo precisa do valor que se vê: um título cujo estilo
 * manda "manter com o próximo" não traz `keepNext` no nó. Sem folha de estilos,
 * os atributos crus.
 */
export function effectiveAttrs(block: StyledBlock, sheet: StyleSheet | null): Record<string, unknown> {
  const attrs: Record<string, unknown> = { ...(block.attrs ?? {}) }
  const style = blockStyleOfNode(block, sheet)
  if (style === null) return attrs

  const inherited = styleAttrsOf(style)
  for (const [name, value] of Object.entries(inherited)) {
    if (attrs[name] === null || attrs[name] === undefined) attrs[name] = value
  }
  return attrs
}

/** O critério da regra do CSS: o id declarado; o título sem id, pelo nome `heading N`; o resto, o padrão. */
export function blockStyleOfNode(block: StyledBlock, sheet: StyleSheet | null): ResolvedStyle | null {
  const type = typeof block.type === 'string' ? block.type : block.type.name
  if (sheet === null || (type !== 'paragraph' && type !== 'heading')) return null

  const attrs = block.attrs ?? {}
  const styleId = typeof attrs['styleId'] === 'string' && attrs['styleId'] !== '' ? attrs['styleId'] : null
  const level = Number(attrs['level'])
  return styleId === null && type === 'heading' && Number.isInteger(level) && level >= 1
    ? headingStyleOf(sheet, level)
    : resolveStyle(sheet, styleId)
}

/** Sem os padrões do documento: por cima, eles apagariam o que o estilo do parágrafo deu ao trecho. */
export function resolveCharacterStyle(sheet: StyleSheet, styleId: string): StyleCharacterFormat {
  let character: StyleCharacterFormat = {}
  for (const style of chainOf(sheet, styleId)) character = overlay(character, style.character)
  return character
}

/** O estilo resolvido como o bloco o diria, campo a campo. Ausente é "ninguém disse". */
export function styleAttrsOf({ paragraph, character }: ResolvedStyle): Record<string, unknown> {
  const attrs: Record<string, unknown> = {
    textAlign: paragraph.textAlign,
    indentMm: paragraph.indentMm,
    indentRightMm: paragraph.indentRightMm,
    firstLineMm: paragraph.firstLineMm,
    // Zero quando a cadeia cala, como a regra do CSS: é o que se vê.
    spaceBefore: paragraph.spaceBefore ?? 0,
    spaceAfter: paragraph.spaceAfter ?? 0,
    lineHeight: lineHeightAttrOf(paragraph.lineSpacing, character.fontFamily ?? null),
    background: paragraph.background,
    keepNext: paragraph.keepNext,
    keepLines: paragraph.keepLines,
    widowControl: paragraph.widowControl,
    fontFamily: character.fontFamily,
    fontSize: character.fontSize,
  }
  for (const key of Object.keys(attrs)) if (attrs[key] === undefined) delete attrs[key]
  return attrs
}

/** O número do CSS já multiplicado pela altura natural da fonte; `exact` e `atLeast` em pontos. */
export function lineHeightAttrOf(spacing: LineSpacing | undefined, family: string | null): string {
  if (spacing === undefined) return cssLineHeightOf(1, family)
  if (spacing.kind === 'multiple') return cssLineHeightOf(usableFactor(spacing.factor), family)
  return `${spacing.pt}pt`
}

/**
 * O múltiplo que vale: fora de `(0,5; 4)` é lixo do arquivo, e vale o simples —
 * o mesmo corte de `BodyReader.LineHeightOf`.
 */
export function usableFactor(factor: number): number {
  return factor > 0.5 && factor < 4 ? factor : 1
}
