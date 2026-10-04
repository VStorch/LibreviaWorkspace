/**
 * A tradução entre o que o bloco guarda e o que o diálogo de parágrafo mostra.
 *
 * - A entrelinha é um campo só, como o CSS a escreve (`normal`, `1.8311`,
 *   `14pt`), e no diálogo são duas perguntas: o tipo e o quanto. O número do CSS
 *   **não** é o fator do Word: já vem multiplicado pela altura natural da fonte
 *   (ver `line-metrics.ts`).
 * - O recuo de primeira linha é um número com sinal (`text-indent` negativo é o
 *   deslocamento do Word); no diálogo, uma escolha e uma medida positiva.
 */

import { INDENT_STEP_MM } from '@services/units.js'
import { cssLineHeightOf, explicitCssLineHeightOf, lineFactorOf } from './line-metrics.js'

export const LineSpacingKind = {
  /** A altura que a própria fonte pede — o silêncio do arquivo. */
  Single: 'single',
  /** Múltiplo da altura natural: 1,15, 1,5, duplo. */
  Multiple: 'multiple',
  /**
   * "Pelo menos" (`w:lineRule="atLeast"`), e não "exatamente": `exact` corta o
   * que não cabe na altura, e perder meia linha é perder conteúdo.
   */
  AtLeast: 'at-least',
} as const

export type LineSpacingKind = (typeof LineSpacingKind)[keyof typeof LineSpacingKind]

export const FirstLineKind = {
  None: 'none',
  Indent: 'indent',
  Hanging: 'hanging',
} as const

export type FirstLineKind = (typeof FirstLineKind)[keyof typeof FirstLineKind]

export const TextAlignment = {
  Left: 'left',
  Center: 'center',
  Right: 'right',
  Justify: 'justify',
} as const

export type TextAlignment = (typeof TextAlignment)[keyof typeof TextAlignment]

/** Já em números, nunca em texto de `<input>`. */
export interface ParagraphDraft {
  readonly align: TextAlignment
  /** Pontos, como o Word os mostra. */
  readonly spaceBefore: number
  readonly spaceAfter: number
  readonly lineSpacingKind: LineSpacingKind
  /** Fator quando o tipo é múltiplo; pontos quando é "pelo menos". */
  readonly lineSpacingValue: number
  /** Milímetros, como a régua da configuração de página. */
  readonly indentLeftMm: number
  readonly indentRightMm: number
  readonly firstLineKind: FirstLineKind
  readonly firstLineMm: number
  readonly keepNext: boolean
  readonly keepLines: boolean
  /** Viúvas e órfãs: ligado quando nada diz o contrário, como no Word. */
  readonly widowControl: boolean
}

export const MAX_SPACING_PT = 600
export const MAX_INDENT_MM = 200
/**
 * O gravador recusa fator fora de `(0,5; 4)` e registra perda (ver
 * `ParagraphFormat.ApplyLineHeight`); o diálogo não oferece o que ele não escreve.
 */
export const MIN_LINE_FACTOR = 0.51
export const MAX_LINE_FACTOR = 3.99

export const DEFAULT_PARAGRAPH_DRAFT: ParagraphDraft = {
  align: TextAlignment.Left,
  spaceBefore: 0,
  spaceAfter: 0,
  lineSpacingKind: LineSpacingKind.Single,
  lineSpacingValue: 1.15,
  indentLeftMm: 0,
  indentRightMm: 0,
  firstLineKind: FirstLineKind.None,
  firstLineMm: 0,
  keepNext: false,
  keepLines: false,
  widowControl: true,
}

/**
 * Valor nulo apaga o atributo. O diálogo escreve só o que a pessoa mudou em
 * relação ao que se vê: transformar herdado em direto desligaria o bloco do
 * estilo sem ninguém pedir.
 */
export interface ParagraphAttrs {
  readonly textAlign: TextAlignment | null
  readonly spaceBefore: number | null
  readonly spaceAfter: number | null
  readonly lineHeight: string | null
  readonly indentMm: number | null
  readonly indentRightMm: number | null
  readonly firstLineMm: number | null
  readonly keepNext: boolean | null
  readonly keepLines: boolean | null
  readonly widowControl: boolean | null
  /**
   * O recuo em passos de `Ctrl+]`. O gravador soma medida e nível; quando o
   * recuo muda pelo diálogo, o nível é zerado para não somar duas vezes.
   */
  readonly indent: number
}

const clamp = (value: number, min: number, max: number): number =>
  Number.isFinite(value) ? Math.min(Math.max(value, min), max) : min

/** A precisão com que o leitor entrega a medida. */
const round = (value: number): number => Math.round(value * 10) / 10

function numberOf(value: unknown): number | null {
  if (value === null || value === undefined || value === '') return null
  const parsed = Number(value)
  return Number.isFinite(parsed) ? parsed : null
}

/** Atributos crus, porque um bloco importado traz medidas que nenhum campo do diálogo criou. */
export function paragraphDraftFrom(attrs: Record<string, unknown>): ParagraphDraft {
  const firstLine = numberOf(attrs['firstLineMm']) ?? 0
  const indentLevel = numberOf(attrs['indent']) ?? 0

  return {
    align: alignmentOf(attrs['textAlign']),
    spaceBefore: clamp(round(numberOf(attrs['spaceBefore']) ?? 0), 0, MAX_SPACING_PT),
    spaceAfter: clamp(round(numberOf(attrs['spaceAfter']) ?? 0), 0, MAX_SPACING_PT),
    ...lineSpacingOf(attrs['lineHeight'], fontStackOf(attrs)),
    // O nível de `Ctrl+]` entra como medida, para o campo mostrar o recuo que se vê.
    indentLeftMm: clamp(
      round((numberOf(attrs['indentMm']) ?? 0) + indentLevel * INDENT_STEP_MM),
      0,
      MAX_INDENT_MM,
    ),
    indentRightMm: clamp(round(numberOf(attrs['indentRightMm']) ?? 0), 0, MAX_INDENT_MM),
    firstLineKind:
      firstLine > 0 ? FirstLineKind.Indent : firstLine < 0 ? FirstLineKind.Hanging : FirstLineKind.None,
    firstLineMm: clamp(round(Math.abs(firstLine)), 0, MAX_INDENT_MM),
    keepNext: attrs['keepNext'] === true,
    keepLines: attrs['keepLines'] === true,
    widowControl: attrs['widowControl'] !== false,
  }
}

function alignmentOf(value: unknown): TextAlignment {
  const found = Object.values(TextAlignment).find((align) => align === value)
  return found ?? TextAlignment.Left
}

function fontStackOf(attrs: Record<string, unknown>): string | null {
  const stack = attrs['fontFamily']
  return typeof stack === 'string' && stack.trim() !== '' ? stack : null
}

function lineSpacingOf(
  value: unknown,
  fontStack: string | null,
): Pick<ParagraphDraft, 'lineSpacingKind' | 'lineSpacingValue'> {
  if (typeof value !== 'string' || value === '' || value.toLowerCase() === 'normal') {
    return {
      lineSpacingKind: LineSpacingKind.Single,
      lineSpacingValue: DEFAULT_PARAGRAPH_DRAFT.lineSpacingValue,
    }
  }

  if (value.toLowerCase().endsWith('pt')) {
    const points = numberOf(value.slice(0, -2))
    return points === null || points <= 0
      ? {
          lineSpacingKind: LineSpacingKind.Single,
          lineSpacingValue: DEFAULT_PARAGRAPH_DRAFT.lineSpacingValue,
        }
      : {
          lineSpacingKind: LineSpacingKind.AtLeast,
          lineSpacingValue: clamp(round(points), 1, MAX_SPACING_PT),
        }
  }

  const css = numberOf(value)
  // O número do atributo é medida de CSS, e o do diálogo é linha do Word: um
  // parágrafo de Calibri sem entrelinha declarada chega como 1,2207 e é "Simples".
  const factor = css === null || css <= 0 ? 1 : lineFactorOf(css, fontStack)

  return factor === 1
    ? {
        lineSpacingKind: LineSpacingKind.Single,
        lineSpacingValue: DEFAULT_PARAGRAPH_DRAFT.lineSpacingValue,
      }
    : {
        lineSpacingKind: LineSpacingKind.Multiple,
        lineSpacingValue: clamp(factor, MIN_LINE_FACTOR, MAX_LINE_FACTOR),
      }
}

/** O deslocamento vira recuo negativo, como o `w:hanging` do Word. */
export function signedFirstLineMm(draft: ParagraphDraft): number {
  if (draft.firstLineKind === FirstLineKind.None) return 0
  return draft.firstLineKind === FirstLineKind.Hanging
    ? -Math.abs(draft.firstLineMm)
    : Math.abs(draft.firstLineMm)
}

/**
 * `effective` é o que o bloco vale com o estilo por baixo (`effectiveAttrs`), e
 * é dele que o formulário nasceu. Grupo de campos igual ao da abertura devolve o
 * atributo cru intocado; diferente vira formatação direta. Zero só vira atributo
 * quando desfaz algo do estilo.
 */
export function paragraphAttrsFrom(
  draft: ParagraphDraft,
  attrs: Record<string, unknown> = {},
  effective: Record<string, unknown> = attrs,
): ParagraphAttrs {
  const shown = paragraphDraftFrom(effective)
  const kept = (name: string): unknown => attrs[name] ?? null
  const same = <K extends keyof ParagraphDraft>(...keys: K[]): boolean =>
    keys.every((key) => draft[key] === shown[key])

  const firstLine = signedFirstLineMm(draft)

  const indentSame = same('indentLeftMm')

  return {
    textAlign: same('align')
      ? (kept('textAlign') as TextAlignment | null)
      : keptAlignment(draft.align, effective),
    spaceBefore: same('spaceBefore')
      ? (kept('spaceBefore') as number | null)
      : clamp(round(draft.spaceBefore), 0, MAX_SPACING_PT),
    spaceAfter: same('spaceAfter')
      ? (kept('spaceAfter') as number | null)
      : clamp(round(draft.spaceAfter), 0, MAX_SPACING_PT),
    lineHeight: same('lineSpacingKind', 'lineSpacingValue')
      ? (kept('lineHeight') as string | null)
      : lineHeightOf(draft, effective),
    indentMm: indentSame
      ? (kept('indentMm') as number | null)
      : measureOf(draft.indentLeftMm, effective, 'indentMm', 0),
    indentRightMm: same('indentRightMm')
      ? (kept('indentRightMm') as number | null)
      : measureOf(draft.indentRightMm, effective, 'indentRightMm', 0),
    firstLineMm: same('firstLineKind', 'firstLineMm')
      ? (kept('firstLineMm') as number | null)
      : measureOf(firstLine, effective, 'firstLineMm', -MAX_INDENT_MM),
    keepNext: same('keepNext') ? (kept('keepNext') as boolean | null) : draft.keepNext ? true : false,
    keepLines: same('keepLines') ? (kept('keepLines') as boolean | null) : draft.keepLines ? true : false,
    widowControl: same('widowControl')
      ? (kept('widowControl') as boolean | null)
      : draft.widowControl
        ? true
        : false,
    indent: indentSame ? (numberOf(attrs['indent']) ?? 0) : 0,
  }
}

/** Zero é ausência, a não ser que o estilo recue: aí é um zero explícito que o desfaz. */
function measureOf(
  value: number,
  effective: Record<string, unknown>,
  name: string,
  min: number,
): number | null {
  if (value !== 0) return clamp(round(value), min, MAX_INDENT_MM)
  const inherited = numberOf(effective[name]) ?? 0
  return inherited === 0 ? null : 0
}

/** "À esquerda" só é escrito quando algo declara outro alinhamento: apagar deixaria o estilo voltar. */
function keptAlignment(align: TextAlignment, effective: Record<string, unknown>): TextAlignment | null {
  const declared = typeof effective['textAlign'] === 'string' ? effective['textAlign'] : null
  return align === TextAlignment.Left && declared === null ? null : align
}

/** A fonte dá a altura natural da linha, e o que já estava escrito decide o espaçamento simples. */
function lineHeightOf(draft: ParagraphDraft, attrs: Record<string, unknown>): string {
  if (draft.lineSpacingKind === LineSpacingKind.AtLeast) {
    return `${clamp(round(draft.lineSpacingValue), 1, MAX_SPACING_PT)}pt`
  }

  const fontStack = fontStackOf(attrs)

  if (draft.lineSpacingKind === LineSpacingKind.Single) {
    const current = attrs['lineHeight']

    // O simples do arquivo é o silêncio dele: trocar por número reescreveria o
    // bloco sem mudar nada do que se vê.
    if (current === null || current === undefined || current === '') return 'normal'
    if (typeof current === 'string' && current.toLowerCase() === 'normal') return 'normal'

    // Já era o fator 1, na forma do leitor: fica como está.
    const css = numberOf(current)
    if (css !== null && css > 0 && lineFactorOf(css, fontStack) === 1) return String(current)

    // Havia medida declarada: `normal` não serve, porque o gravador não o grava
    // e o `w:line` antigo ficaria de pé.
    return explicitCssLineHeightOf(1, fontStack)
  }

  return cssLineHeightOf(clamp(draft.lineSpacingValue, MIN_LINE_FACTOR, MAX_LINE_FACTOR), fontStack)
}

/** Para a barra e os atalhos `Ctrl+1`, `Ctrl+5` e `Ctrl+2`, que falam em linha, como o Word: `''` é simples. */
export function lineSpacingChoiceOf(attrs: Record<string, unknown>): string {
  const spacing = lineSpacingOf(attrs['lineHeight'], fontStackOf(attrs))

  if (spacing.lineSpacingKind === LineSpacingKind.Single) return ''
  if (spacing.lineSpacingKind === LineSpacingKind.AtLeast) return `${spacing.lineSpacingValue}pt`
  return String(spacing.lineSpacingValue)
}

export function lineHeightAttrFrom(choice: string, attrs: Record<string, unknown>): string {
  const spacing = choice.endsWith('pt')
    ? { lineSpacingKind: LineSpacingKind.AtLeast, lineSpacingValue: numberOf(choice.slice(0, -2)) ?? 12 }
    : choice.trim() === ''
      ? { lineSpacingKind: LineSpacingKind.Single, lineSpacingValue: 1 }
      : { lineSpacingKind: LineSpacingKind.Multiple, lineSpacingValue: numberOf(choice) ?? 1 }

  return lineHeightOf({ ...DEFAULT_PARAGRAPH_DRAFT, ...spacing }, attrs)
}

/** Só os números podem sair de faixa: o resto é escolha fechada. */
export function isValidParagraphDraft(draft: ParagraphDraft): boolean {
  const inRange = (value: number, min: number, max: number): boolean =>
    Number.isFinite(value) && value >= min && value <= max

  if (!inRange(draft.spaceBefore, 0, MAX_SPACING_PT)) return false
  if (!inRange(draft.spaceAfter, 0, MAX_SPACING_PT)) return false
  if (!inRange(draft.indentLeftMm, 0, MAX_INDENT_MM)) return false
  if (!inRange(draft.indentRightMm, 0, MAX_INDENT_MM)) return false
  if (!inRange(draft.firstLineMm, 0, MAX_INDENT_MM)) return false

  if (draft.lineSpacingKind === LineSpacingKind.Multiple) {
    return inRange(draft.lineSpacingValue, MIN_LINE_FACTOR, MAX_LINE_FACTOR)
  }
  if (draft.lineSpacingKind === LineSpacingKind.AtLeast) {
    return inRange(draft.lineSpacingValue, 1, MAX_SPACING_PT)
  }

  return true
}
