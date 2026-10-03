/**
 * O múltiplo do OOXML (`w:line` com `w:lineRule="auto"`) é medido sobre a
 * **altura natural da fonte** — `(ascender - descender + lineGap) / unitsPerEm`
 * da tabela `hhea` —, como no Word, no LibreOffice e no `line-height: normal`.
 * Por isso o atributo `lineHeight` guarda o número do CSS, já multiplicado: o
 * leitor multiplica, o gravador divide, e a interface converte para mostrar.
 *
 * `line-metrics.test.ts` compara esta tabela com `LineMetrics.cs`. Só as fontes
 * do instalador e as que elas substituem: para o resto não há palpite honesto.
 */

/** A fonte do editor quando o documento não diz outra. */
export const DEFAULT_NATURAL_LINE_HEIGHT = 1.1499

const LIBERATION_SERIF = DEFAULT_NATURAL_LINE_HEIGHT

/** Nome da família → altura natural. As chaves comparam sem caixa nem espaço. */
export const NATURAL_LINE_HEIGHTS: Readonly<Record<string, number>> = {
  arial: 1.1499,
  helvetica: 1.1499,
  'liberation sans': 1.1499,
  'times new roman': LIBERATION_SERIF,
  'liberation serif': LIBERATION_SERIF,
  'courier new': 1.1328,
  'liberation mono': 1.1328,
  calibri: 1.2207,
  carlito: 1.2207,
  cambria: 1.15,
  caladea: 1.15,
}

/** A altura é a da primeira; a genérica só vale se ela faltar. Espelho de `ParagraphFormat.FirstFont`. */
export function firstFontOf(stack: string | null | undefined): string | null {
  if (typeof stack !== 'string') return null
  const first =
    stack
      .split(',')[0]
      ?.trim()
      .replace(/^['"]|['"]$/g, '') ?? ''
  return first.length > 0 ? first : null
}

/** Espelho de `LineMetrics.Of`: pilha vazia é a fonte do editor; fonte desconhecida é `null`. */
export function naturalLineHeightOf(stack: string | null | undefined): number | null {
  const first = firstFontOf(stack)
  if (first === null) return DEFAULT_NATURAL_LINE_HEIGHT
  return NATURAL_LINE_HEIGHTS[first.toLowerCase()] ?? null
}

/**
 * Espelha `BodyReader.Multiple`: com fonte desconhecida, o fator 1 sai como
 * `normal` e o resto pelo palpite de 1,15, o mesmo do gravador.
 */
export function cssLineHeightOf(factor: number, stack: string | null | undefined): string {
  const natural = naturalLineHeightOf(stack)
  if (natural !== null) return numberText(factor * natural)
  return factor === 1 ? 'normal' : numberText(factor * DEFAULT_NATURAL_LINE_HEIGHT)
}

/** Para sobrescrever uma entrelinha declarada: `normal` o gravador não grava, e o `w:line` antigo ficaria. */
export function explicitCssLineHeightOf(factor: number, stack: string | null | undefined): string {
  return numberText(factor * (naturalLineHeightOf(stack) ?? DEFAULT_NATURAL_LINE_HEIGHT))
}

/** Ao centésimo, a precisão que sobrevive à ida e volta: sem isso "1,5" voltaria 1,4999. */
export function lineFactorOf(css: number, stack: string | null | undefined): number {
  const natural = naturalLineHeightOf(stack) ?? DEFAULT_NATURAL_LINE_HEIGHT
  return Math.round((css / natural) * 100) / 100
}

/** Quatro casas no máximo, sem zero à direita — a forma que o leitor escreve. */
function numberText(value: number): string {
  return String(Math.round(value * 10_000) / 10_000)
}
