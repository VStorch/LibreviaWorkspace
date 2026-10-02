import { firstFontOf, lineFactorOf } from './line-metrics.js'
import type { StyleCharacterFormat, StyleParagraphFormat } from './styles.js'

/**
 * As peças miúdas da exportação para ODT (M11): escapar, medir, traduzir cor e
 * fonte, e o caderno de estilos automáticos.
 *
 * O ODF guarda a formatação em **estilos**, não no texto: cada parágrafo com
 * formatação direta aponta um estilo automático que herda do estilo nomeado
 * dele e declara só a diferença. O caderno (`StyleBook`) dá um nome a cada
 * combinação e reaproveita o nome quando a mesma combinação volta — mil
 * parágrafos centralizados são um estilo só.
 */

/** Os caracteres que o XML 1.0 não aceita nem escapados. */
// eslint-disable-next-line no-control-regex
const INVALID_XML = /[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g

/** Texto de nó ou de atributo, escapado; o que o XML não aceita sai. */
export function xml(text: string): string {
  return text
    .replace(INVALID_XML, '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&apos;')
}

/** ` nome="valor"`, ou nada quando não há valor. */
export function attr(name: string, value: string | number | null | undefined): string {
  if (value === null || value === undefined || value === '') return ''
  return ` ${name}="${xml(String(value))}"`
}

/**
 * O texto de um trecho com o espaço que o ODF guarda: ele junta espaços em
 * sequência e descarta os do começo, então a sequência vira `text:s`, a
 * tabulação `text:tab` e a quebra `text:line-break`. O espaço que abre o trecho
 * também vira `text:s` — o trecho anterior pode ter terminado em espaço, e os
 * dois se juntariam num só.
 */
export function odfText(text: string): string {
  let out = ''
  let index = 0
  while (index < text.length) {
    const char = text[index]!
    if (char === '\t') {
      out += '<text:tab/>'
      index++
    } else if (char === '\n') {
      out += '<text:line-break/>'
      index++
    } else if (char === ' ') {
      let end = index
      while (text[end] === ' ') end++
      const count = end - index
      const leading = index === 0 || text[index - 1] === '\t' || text[index - 1] === '\n'
      if (leading) out += count === 1 ? '<text:s/>' : `<text:s text:c="${count}"/>`
      else out += count === 1 ? ' ' : ` <text:s${count === 2 ? '' : ` text:c="${count - 1}"`}/>`
      index = end
    } else {
      let end = index
      while (end < text.length && !'\t\n '.includes(text[end]!)) end++
      out += xml(text.slice(index, end))
      index = end
    }
  }
  return out
}

/** Milímetros com até três casas — o bastante para um centésimo de ponto. */
export function mm(value: number): string {
  return `${Math.round(value * 1000) / 1000}mm`
}

export function pt(value: number): string {
  return `${Math.round(value * 100) / 100}pt`
}

/** Um número finito, ou `null` — o atributo vem do documento, e pode vir de tudo. */
export function finite(value: unknown): number | null {
  if (value === null || value === undefined || value === '' || typeof value === 'boolean') return null
  const number = Number(value)
  return Number.isFinite(number) ? number : null
}

/** `12pt`, `16px` ou `12` (pontos) → pontos. */
export function pointsOf(value: unknown): number | null {
  if (typeof value === 'number') return Number.isFinite(value) && value > 0 ? value : null
  if (typeof value !== 'string') return null
  const match = /^\s*(\d+(?:\.\d+)?)\s*(pt|px)?\s*$/i.exec(value)
  if (match === null) return null
  const number = Number(match[1])
  if (!(number > 0)) return null
  return match[2]?.toLowerCase() === 'px' ? number * 0.75 : number
}

/** As cores com nome que o leitor pode trazer — as do realce do Word e as básicas do CSS. */
const NAMED_COLORS: Readonly<Record<string, string>> = {
  black: '#000000',
  white: '#ffffff',
  red: '#ff0000',
  green: '#00ff00',
  lime: '#00ff00',
  blue: '#0000ff',
  yellow: '#ffff00',
  cyan: '#00ffff',
  aqua: '#00ffff',
  magenta: '#ff00ff',
  fuchsia: '#ff00ff',
  darkblue: '#000080',
  navy: '#000080',
  darkcyan: '#008080',
  teal: '#008080',
  darkgreen: '#008000',
  darkmagenta: '#800080',
  purple: '#800080',
  darkred: '#800000',
  maroon: '#800000',
  darkyellow: '#808000',
  olive: '#808000',
  darkgray: '#808080',
  gray: '#808080',
  grey: '#808080',
  lightgray: '#c0c0c0',
  silver: '#c0c0c0',
  orange: '#ffa500',
}

/** A cor em `#rrggbb`, que é a única forma que o ODF aceita; `null` para o resto. */
export function odfColor(value: unknown): string | null {
  if (typeof value !== 'string') return null
  const text = value.trim().toLowerCase()
  const hex = /^#?([0-9a-f]{6}|[0-9a-f]{3})$/.exec(text)
  if (hex !== null) {
    const digits = hex[1]!
    return digits.length === 3 ? `#${[...digits].map((digit) => digit + digit).join('')}` : `#${digits}`
  }
  const rgb = /^rgba?\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})\s*(?:,\s*([\d.]+)\s*)?\)$/.exec(text)
  if (rgb !== null) {
    if (rgb[4] !== undefined && Number(rgb[4]) === 0) return null
    return `#${[rgb[1], rgb[2], rgb[3]]
      .map((channel) => Math.min(255, Number(channel)).toString(16).padStart(2, '0'))
      .join('')}`
  }
  return NAMED_COLORS[text.replace(/[\s_-]/g, '')] ?? null
}

/** As fontes que os estilos citam: cada uma ganha a declaração dela no arquivo. */
export class FontBook {
  private readonly names = new Set<string>()

  /** O nome da fonte na pilha de CSS (`Calibri, Carlito, sans-serif` → `Calibri`). */
  use(stack: unknown): string | null {
    if (typeof stack !== 'string') return null
    const name = firstFontOf(stack)
    if (name === null || name === '') return null
    this.names.add(name)
    return name
  }

  xml(): string {
    const faces = [...this.names]
      .sort()
      .map((name) => {
        const family = /\s/.test(name) ? `'${name}'` : name
        return `<style:font-face${attr('style:name', name)}${attr('svg:font-family', family)}/>`
      })
      .join('')
    return `<office:font-face-decls>${faces}</office:font-face-decls>`
  }
}

/** O que um trecho de texto diz, já com o vocabulário do ODF em mente. */
export interface CharacterProps {
  fontFamily?: unknown
  fontSize?: unknown
  bold?: boolean | undefined
  italic?: boolean | undefined
  underline?: boolean | undefined
  strike?: boolean | undefined
  caps?: boolean | undefined
  smallCaps?: boolean | undefined
  /** `super` ou `sub`. */
  position?: string | undefined
  color?: unknown
  background?: unknown
  mono?: boolean | undefined
}

/** Os atributos de `style:text-properties` — vazio quando o trecho não diz nada. */
export function textProperties(props: CharacterProps, fonts: FontBook): string {
  const parts: string[] = []
  const font = props.mono === true ? 'Liberation Mono' : fonts.use(props.fontFamily)
  if (props.mono === true) fonts.use('Liberation Mono')
  if (font !== null) {
    parts.push(attr('style:font-name', font), attr('style:font-name-complex', font))
  }
  const size = pointsOf(props.fontSize)
  if (size !== null) {
    parts.push(
      attr('fo:font-size', pt(size)),
      attr('style:font-size-asian', pt(size)),
      attr('style:font-size-complex', pt(size)),
    )
  }
  if (props.bold !== undefined) {
    const weight = props.bold ? 'bold' : 'normal'
    parts.push(
      attr('fo:font-weight', weight),
      attr('style:font-weight-asian', weight),
      attr('style:font-weight-complex', weight),
    )
  }
  if (props.italic !== undefined) {
    const style = props.italic ? 'italic' : 'normal'
    parts.push(
      attr('fo:font-style', style),
      attr('style:font-style-asian', style),
      attr('style:font-style-complex', style),
    )
  }
  if (props.underline !== undefined) {
    parts.push(
      props.underline
        ? ' style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"'
        : ' style:text-underline-style="none"',
    )
  }
  if (props.strike !== undefined) {
    parts.push(
      props.strike
        ? ' style:text-line-through-style="solid" style:text-line-through-type="single"'
        : ' style:text-line-through-style="none"',
    )
  }
  if (props.caps === true) parts.push(' fo:text-transform="uppercase"')
  if (props.smallCaps === true) parts.push(' fo:font-variant="small-caps"')
  if (props.position === 'super') parts.push(' style:text-position="super 58%"')
  if (props.position === 'sub') parts.push(' style:text-position="sub 58%"')
  const color = odfColor(props.color)
  if (color !== null) parts.push(attr('fo:color', color))
  const background = odfColor(props.background)
  if (background !== null) parts.push(attr('fo:background-color', background))
  return parts.join('')
}

/** O que um estilo de caractere (ou a parte de texto de um de parágrafo) diz. */
export function characterPropsOfStyle(format: StyleCharacterFormat | undefined): CharacterProps {
  if (format === undefined) return {}
  const vertical = format.verticalAlign?.toLowerCase()
  return {
    fontFamily: format.fontFamily,
    fontSize: format.fontSize,
    bold: format.bold,
    italic: format.italic,
    underline: format.underline,
    strike: format.strike,
    caps: format.allCaps,
    smallCaps: format.smallCaps,
    position: vertical?.startsWith('super') ? 'super' : vertical?.startsWith('sub') ? 'sub' : undefined,
    color: format.color,
    background: format.highlight,
  }
}

/** O alinhamento do documento no vocabulário do ODF. */
export function odfAlign(value: unknown): string | null {
  switch (typeof value === 'string' ? value.toLowerCase() : '') {
    case 'left':
    case 'start':
      return 'start'
    case 'right':
    case 'end':
      return 'end'
    case 'center':
      return 'center'
    case 'justify':
    case 'both':
    case 'distribute':
      return 'justify'
    default:
      return null
  }
}

/** O que um parágrafo diz, nas unidades do documento (mm e pt). */
export interface ParagraphProps {
  align?: unknown
  marginLeftMm?: number | null
  marginRightMm?: number | null
  textIndentMm?: number | null
  spaceBeforePt?: number | null
  spaceAfterPt?: number | null
  /** Múltiplo da altura natural da linha. */
  lineFactor?: number | null
  /** Altura mínima em pontos — o `exact` e o `atLeast` do arquivo. */
  lineAtLeastPt?: number | null
  keepNext?: boolean | undefined
  keepLines?: boolean | undefined
  widowControl?: boolean | undefined
  breakBefore?: 'page' | 'column' | null
  breakAfter?: 'page' | 'column' | null
  contextualSpacing?: boolean | undefined
  background?: unknown
  /** O número com que a página recomeça, com a página mestra da seção. */
  pageNumber?: number | 'auto' | null
  borderBottom?: string | null
}

const isSet = <T>(value: T | null | undefined): value is T => value !== null && value !== undefined

/** Os atributos de `style:paragraph-properties` — vazio quando o parágrafo não diz nada. */
export function paragraphProperties(props: ParagraphProps): string {
  const parts: string[] = []
  const align = odfAlign(props.align)
  if (align !== null) parts.push(attr('fo:text-align', align))
  if (isSet(props.marginLeftMm)) parts.push(attr('fo:margin-left', mm(props.marginLeftMm)))
  if (isSet(props.marginRightMm)) parts.push(attr('fo:margin-right', mm(props.marginRightMm)))
  if (isSet(props.textIndentMm)) parts.push(attr('fo:text-indent', mm(props.textIndentMm)))
  if (isSet(props.spaceBeforePt)) parts.push(attr('fo:margin-top', pt(props.spaceBeforePt)))
  if (isSet(props.spaceAfterPt)) parts.push(attr('fo:margin-bottom', pt(props.spaceAfterPt)))
  if (isSet(props.lineAtLeastPt)) parts.push(attr('style:line-height-at-least', pt(props.lineAtLeastPt)))
  else if (isSet(props.lineFactor))
    parts.push(attr('fo:line-height', `${Math.round(props.lineFactor * 100)}%`))
  if (props.keepNext !== undefined) parts.push(attr('fo:keep-with-next', props.keepNext ? 'always' : 'auto'))
  if (props.keepLines !== undefined) parts.push(attr('fo:keep-together', props.keepLines ? 'always' : 'auto'))
  if (props.widowControl !== undefined) {
    const lines = props.widowControl ? '2' : '0'
    parts.push(attr('fo:widows', lines), attr('fo:orphans', lines))
  }
  if (isSet(props.breakBefore)) parts.push(attr('fo:break-before', props.breakBefore))
  if (isSet(props.breakAfter)) parts.push(attr('fo:break-after', props.breakAfter))
  if (props.contextualSpacing !== undefined) {
    parts.push(attr('style:contextual-spacing', props.contextualSpacing ? 'true' : 'false'))
  }
  const background = odfColor(props.background)
  if (background !== null) parts.push(attr('fo:background-color', background))
  if (isSet(props.pageNumber)) parts.push(attr('style:page-number', props.pageNumber))
  if (isSet(props.borderBottom)) {
    parts.push(attr('fo:border-bottom', props.borderBottom), attr('fo:padding-bottom', '0.5mm'))
  }
  return parts.join('')
}

/** O que um estilo de parágrafo diz, traduzido. */
export function paragraphPropsOfStyle(format: StyleParagraphFormat | undefined): ParagraphProps {
  if (format === undefined) return {}
  const spacing = format.lineSpacing
  return {
    align: format.textAlign,
    marginLeftMm: format.indentMm ?? null,
    marginRightMm: format.indentRightMm ?? null,
    textIndentMm: format.firstLineMm ?? null,
    spaceBeforePt: format.spaceBefore ?? null,
    spaceAfterPt: format.spaceAfter ?? null,
    lineFactor: spacing?.kind === 'multiple' ? spacing.factor : null,
    lineAtLeastPt: spacing !== undefined && spacing.kind !== 'multiple' ? spacing.pt : null,
    keepNext: format.keepNext,
    keepLines: format.keepLines,
    widowControl: format.widowControl,
    breakBefore: format.pageBreakBefore === true ? 'page' : null,
    contextualSpacing: format.contextualSpacing,
    background: format.background,
  }
}

/**
 * A entrelinha do atributo do bloco (`1.2422`, `normal`, `12pt`) como o ODF a
 * quer: o múltiplo da altura natural — o número do CSS dividido por ela — ou a
 * altura mínima em pontos.
 */
export function lineSpacingOfAttr(
  value: unknown,
  fontFamily: unknown,
): Pick<ParagraphProps, 'lineFactor' | 'lineAtLeastPt'> {
  if (typeof value === 'string' && /pt\s*$/i.test(value)) return { lineAtLeastPt: pointsOf(value) }
  if (value === 'normal') return { lineFactor: 1 }
  const css = finite(value)
  if (css === null || css <= 0) return {}
  return { lineFactor: lineFactorOf(css, typeof fontFamily === 'string' ? fontFamily : null) }
}

/**
 * Os estilos automáticos de um arquivo do pacote (o `content.xml` ou o
 * `styles.xml`, que não enxergam os automáticos um do outro — daí o prefixo).
 */
export class StyleBook {
  private readonly names = new Map<string, string>()
  private readonly parts: string[] = []
  private readonly counts = new Map<string, number>()

  constructor(private readonly prefix: string) {}

  /** O nome do estilo com estes atributos e este corpo — o mesmo para a mesma combinação. */
  style(family: string, letter: string, attrs: string, body: string): string {
    return this.add(
      letter,
      `${family}|${attrs}|${body}`,
      (name) =>
        `<style:style${attr('style:name', name)}${attr('style:family', family)}${attrs}>${body}</style:style>`,
    )
  }

  /** Um estilo de lista, que é outro elemento. */
  list(body: string): string {
    return this.add(
      'L',
      `list|${body}`,
      (name) => `<text:list-style${attr('style:name', name)}>${body}</text:list-style>`,
    )
  }

  private add(letter: string, key: string, build: (name: string) => string): string {
    const known = this.names.get(key)
    if (known !== undefined) return known
    const count = (this.counts.get(letter) ?? 0) + 1
    this.counts.set(letter, count)
    const name = `${this.prefix}${letter}${count}`
    this.names.set(key, name)
    this.parts.push(build(name))
    return name
  }

  xml(): string {
    return this.parts.join('')
  }
}
