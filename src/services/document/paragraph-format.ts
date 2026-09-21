/**
 * A formatação de parágrafo como o diálogo a mostra, e como o bloco a guarda.
 *
 * Os atributos já existem no modelo e o gravador já sabe escrevê-los (ver
 * `extensions/block-format.ts` e `ParagraphFormat.cs`): o que faltava era a
 * tradução entre o que o arquivo representa e o que uma pessoa consegue
 * preencher num formulário. As duas formas não são a mesma:
 *
 *  - a entrelinha mora num campo só, como o CSS a escreve — `normal`, um número
 *    (`1.8311`) ou uma medida (`14pt`) —, e no diálogo são duas perguntas: que
 *    tipo de espaçamento, e quanto. E o número do CSS **não** é o fator do Word:
 *    é ele já multiplicado pela altura natural da fonte (ver `line-metrics.ts`).
 *    Tratar um pelo outro é o que fazia todo parágrafo de Calibri abrir como
 *    "Múltiplo 1,2207" e "1,5" gravar 1,23 linha no arquivo;
 *  - o recuo de primeira linha é **um** número com sinal, porque no CSS
 *    `text-indent` negativo é o deslocamento do Word; no diálogo são a escolha
 *    ("nenhum", "primeira linha", "deslocamento") e uma medida positiva.
 *
 * Aqui só a conversão e os limites. Nada de React, nada de Tiptap: é o que
 * permite testar as duas direções sem montar editor nenhum.
 */

import { cssLineHeightOf, explicitCssLineHeightOf, lineFactorOf } from './line-metrics.js'

/** Como a entrelinha é medida. Os três casos que o gravador sabe escrever. */
export const LineSpacingKind = {
  /** A altura que a própria fonte pede — o silêncio do arquivo. */
  Single: 'single',
  /** Múltiplo da altura natural: 1,15, 1,5, duplo. */
  Multiple: 'multiple',
  /**
   * Medida fixa em pontos.
   *
   * "Pelo menos", e não "exatamente": é assim que o gravador a escreve
   * (`w:lineRule="atLeast"`), porque `exact` corta o que não couber na altura
   * declarada — e perder metade de uma linha é perda de conteúdo.
   */
  AtLeast: 'at-least',
} as const

export type LineSpacingKind = (typeof LineSpacingKind)[keyof typeof LineSpacingKind]

/** O que o recuo da primeira linha faz: nada, entra ou sai. */
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

/** O formulário inteiro, já em números — nunca em texto de `<input>`. */
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
}

/** Meia polegada, em milímetros: o passo de recuo do OOXML (720 twips). */
export const INDENT_STEP_MM = 12.7

export const MAX_SPACING_PT = 600
export const MAX_INDENT_MM = 200
/**
 * Os limites do fator, e por que são estes.
 *
 * O gravador recusa fator fora de `(0,5; 4)` e registra perda — ver
 * `ParagraphFormat.ApplyLineHeight`. Oferecer no diálogo o que ele não escreve
 * seria prometer o que não se cumpre.
 *
 * São limites do fator **do Word**, e é a conversão de `line-metrics.ts` que os
 * torna verdadeiros: enquanto o número do diálogo ia cru para o atributo, 0,51
 * em Calibri chegava ao gravador como 0,42 — fora da faixa, perda no inventário
 * e nada escrito no arquivo.
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
}

/** Os atributos que o diálogo escreve no bloco. Valor nulo apaga o atributo. */
export interface ParagraphAttrs {
  /**
   * Nulo quando ninguém escolheu alinhamento.
   *
   * Pela mesma razão dos recuos: o leitor omite o atributo no parágrafo sem
   * `w:jc`, e escrever `left` ali acrescentaria um `w:jc` que o documento não
   * tinha — bloco reescrito à custa de um "Aplicar" que não mudou nada.
   */
  readonly textAlign: TextAlignment | null
  readonly spaceBefore: number
  readonly spaceAfter: number
  readonly lineHeight: string
  readonly indentMm: number | null
  readonly indentRightMm: number | null
  readonly firstLineMm: number | null
  readonly keepNext: true | null
  /**
   * O recuo em passos de `Ctrl+]` é zerado.
   *
   * As duas origens somam no gravador (medida do arquivo + nível do editor), e
   * o diálogo fala em milímetros: deixar o nível de pé faria o campo mostrar
   * 10 mm e o arquivo receber 10 mm mais dois passos.
   */
  readonly indent: 0
}

const clamp = (value: number, min: number, max: number): number =>
  Number.isFinite(value) ? Math.min(Math.max(value, min), max) : min

/** Arredonda para o décimo: é a precisão com que o leitor entrega a medida. */
const round = (value: number): number => Math.round(value * 10) / 10

function numberOf(value: unknown): number | null {
  if (value === null || value === undefined || value === '') return null
  const parsed = Number(value)
  return Number.isFinite(parsed) ? parsed : null
}

/**
 * O bloco selecionado → o formulário.
 *
 * Recebe os atributos crus do nó porque é assim que o Tiptap os entrega, e
 * porque um bloco importado traz medidas que nenhum campo do diálogo criou.
 */
export function paragraphDraftFrom(attrs: Record<string, unknown>): ParagraphDraft {
  const firstLine = numberOf(attrs['firstLineMm']) ?? 0
  const indentLevel = numberOf(attrs['indent']) ?? 0

  return {
    align: alignmentOf(attrs['textAlign']),
    spaceBefore: clamp(round(numberOf(attrs['spaceBefore']) ?? 0), 0, MAX_SPACING_PT),
    spaceAfter: clamp(round(numberOf(attrs['spaceAfter']) ?? 0), 0, MAX_SPACING_PT),
    ...lineSpacingOf(attrs['lineHeight'], fontStackOf(attrs)),
    // O nível de `Ctrl+]` entra como medida, senão o recuo que a pessoa vê na
    // tela não é o que o campo mostra. 12,7 mm é o passo de meia polegada que o
    // OOXML usa (720 twips), o mesmo de `TwipsPerIndentLevel` no gravador.
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
  }
}

function alignmentOf(value: unknown): TextAlignment {
  const found = Object.values(TextAlignment).find((align) => align === value)
  return found ?? TextAlignment.Left
}

/** A pilha de fontes do bloco, que é quem diz a altura natural da linha. */
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
  // O número do atributo é medida de CSS; o do diálogo é linha do Word. Um
  // parágrafo de Calibri sem entrelinha declarada chega aqui como 1,2207, e o que
  // a pessoa precisa ver é "Simples".
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

/**
 * O formulário → os atributos do bloco.
 *
 * Zero vira `null` nos recuos de propósito: ausência é o que o leitor produz
 * para o parágrafo sem recuo, e um `0` explícito faria a impressão digital ver
 * diferença onde não há — todo bloco seria reescrito ao salvar.
 */
export function paragraphAttrsFrom(
  draft: ParagraphDraft,
  attrs: Record<string, unknown> = {},
): ParagraphAttrs {
  const firstLine =
    draft.firstLineKind === FirstLineKind.None
      ? 0
      : draft.firstLineKind === FirstLineKind.Hanging
        ? -Math.abs(draft.firstLineMm)
        : Math.abs(draft.firstLineMm)

  return {
    textAlign: keptAlignment(draft.align, attrs),
    spaceBefore: clamp(round(draft.spaceBefore), 0, MAX_SPACING_PT),
    spaceAfter: clamp(round(draft.spaceAfter), 0, MAX_SPACING_PT),
    lineHeight: lineHeightOf(draft, attrs),
    indentMm: draft.indentLeftMm > 0 ? clamp(round(draft.indentLeftMm), 0, MAX_INDENT_MM) : null,
    indentRightMm: draft.indentRightMm > 0 ? clamp(round(draft.indentRightMm), 0, MAX_INDENT_MM) : null,
    firstLineMm: firstLine === 0 ? null : clamp(round(firstLine), -MAX_INDENT_MM, MAX_INDENT_MM),
    keepNext: draft.keepNext ? true : null,
    indent: 0,
  }
}

/**
 * O alinhamento, escrito só quando é escolha.
 *
 * "À esquerda" é o que o formulário mostra para o parágrafo que não declara
 * alinhamento nenhum — então, se era isso que estava lá, continua ausente. Num
 * parágrafo que **declara** outro alinhamento, escolher esquerda é uma decisão, e
 * apagar o atributo deixaria o estilo justificar de volta.
 */
function keptAlignment(align: TextAlignment, attrs: Record<string, unknown>): TextAlignment | null {
  const declared = typeof attrs['textAlign'] === 'string' ? attrs['textAlign'] : null
  return align === TextAlignment.Left && declared === null ? null : align
}

/**
 * A entrelinha como o CSS a escreve — a mesma forma que o leitor produz.
 *
 * Recebe os atributos atuais porque a conversão depende de duas coisas do bloco:
 * a fonte, que dá a altura natural da linha, e o que já estava escrito ali. O
 * segundo importa só no espaçamento simples, e os três caminhos estão comentados
 * abaixo.
 */
function lineHeightOf(draft: ParagraphDraft, attrs: Record<string, unknown>): string {
  if (draft.lineSpacingKind === LineSpacingKind.AtLeast) {
    return `${clamp(round(draft.lineSpacingValue), 1, MAX_SPACING_PT)}pt`
  }

  const fontStack = fontStackOf(attrs)

  if (draft.lineSpacingKind === LineSpacingKind.Single) {
    const current = attrs['lineHeight']

    // Nada declarado, ou `normal`: o simples do arquivo é o silêncio dele, e quem
    // mede melhor aí é o navegador. Trocar por número reescreveria o bloco sem
    // mudar uma linha do que se vê.
    if (current === null || current === undefined || current === '') return 'normal'
    if (typeof current === 'string' && current.toLowerCase() === 'normal') return 'normal'

    // Já era o fator 1, na forma em que o leitor o escreveu: fica como está, pelo
    // mesmo motivo.
    const css = numberOf(current)
    if (css !== null && css > 0 && lineFactorOf(css, fontStack) === 1) return String(current)

    // Havia medida declarada, e agora a escolha é simples. Aqui **não** vale
    // `normal`: é o único valor que o gravador não grava, e o `w:line` antigo
    // ficaria de pé — a entrelinha da pessoa perdida em silêncio.
    return explicitCssLineHeightOf(1, fontStack)
  }

  return cssLineHeightOf(clamp(draft.lineSpacingValue, MIN_LINE_FACTOR, MAX_LINE_FACTOR), fontStack)
}

/**
 * A entrelinha do bloco como os seletores rápidos a falam: `''` para simples, o
 * fator do Word (`1.5`) ou a medida em pontos (`14pt`).
 *
 * Existe para a barra de ferramentas e para os atalhos `Ctrl+1`, `Ctrl+5` e
 * `Ctrl+2`: eles falam em linha, como o Word, e o atributo fala em CSS.
 */
export function lineSpacingChoiceOf(attrs: Record<string, unknown>): string {
  const spacing = lineSpacingOf(attrs['lineHeight'], fontStackOf(attrs))

  if (spacing.lineSpacingKind === LineSpacingKind.Single) return ''
  if (spacing.lineSpacingKind === LineSpacingKind.AtLeast) return `${spacing.lineSpacingValue}pt`
  return String(spacing.lineSpacingValue)
}

/** O inverso: a escolha do seletor rápido → o atributo do bloco. */
export function lineHeightAttrFrom(choice: string, attrs: Record<string, unknown>): string {
  const spacing = choice.endsWith('pt')
    ? { lineSpacingKind: LineSpacingKind.AtLeast, lineSpacingValue: numberOf(choice.slice(0, -2)) ?? 12 }
    : choice.trim() === ''
      ? { lineSpacingKind: LineSpacingKind.Single, lineSpacingValue: 1 }
      : { lineSpacingKind: LineSpacingKind.Multiple, lineSpacingValue: numberOf(choice) ?? 1 }

  return lineHeightOf({ ...DEFAULT_PARAGRAPH_DRAFT, ...spacing }, attrs)
}

/**
 * O formulário está preenchido de um jeito que dá para aplicar?
 *
 * Só o que um número fora de faixa causaria: o resto dos campos é escolha
 * fechada, e o `clamp` já protege a conversão.
 */
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
