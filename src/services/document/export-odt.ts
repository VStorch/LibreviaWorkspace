import { INDENT_STEP_MM, pxToMm } from '@services/units.js'
import { hasBandContent, linesOf, plainBand, type Band, type BandPiece } from './band.js'
import {
  imageData,
  imageExtension,
  isSectionMarkOnly,
  prepareExport,
  safeHref,
  type ExportSource,
} from './export-common.js'
import { fieldKind } from './fields.js'
import type { FloatingObject } from './floating.js'
import { floatsOf } from './floating.js'
import type { ListInfo } from './list-numbering.js'
import {
  contentWidthMm,
  pageDimensionsMm,
  type DocumentModel,
  type DocumentNode,
  type PageSetup,
} from './model.js'
import { mathMlToString, mathText, sanitizeMathMl, type MathElement } from './mathml.js'
import { latexOfEquation } from './mathml-latex.js'
import { NoteKind } from './notes.js'
import {
  attr,
  characterPropsOfStyle,
  finite,
  FontBook,
  lineSpacingOfAttr,
  mm,
  odfColor,
  odfText,
  paragraphProperties,
  paragraphPropsOfStyle,
  StyleBook,
  textProperties,
  xml,
  type CharacterProps,
  type ParagraphProps,
} from './odt-xml.js'
import { blockSections, effectiveSections, sectionBreakOf } from './sections.js'
import { blockStyleOfNode, headingStyleOf } from './style-cascade.js'
import { StyleType, type StyleDefinition, type StyleSheet } from './styles.js'
import { cellBordersFromAttr, type CellBorder } from './table-format.js'
import { zip, type Deflate, type ZipEntry } from './zip.js'

/**
 * Feita aqui, e não no sidecar, porque a cascata de estilos, a conta das listas
 * e a das notas já moram em `@services/document`. Estilos nomeados viram estilos
 * nomeados do ODT; a formatação direta, estilo automático que herda deles. Tudo
 * o que vem do documento é escapado, e só links seguros viram link.
 */

export interface OdtExportOptions {
  /** O programa que gerou o arquivo, para o `meta:generator`. */
  readonly generator?: string
}

type OdtModel = Pick<
  DocumentModel,
  'doc' | 'styles' | 'notes' | 'properties' | 'page' | 'sections' | 'comments'
>

const MIMETYPE = 'application/vnd.oasis.opendocument.text'
const FORMULA_MIMETYPE = 'application/vnd.oasis.opendocument.formula'

const NAMESPACES = [
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"',
  'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"',
  'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"',
  'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"',
  'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"',
  'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"',
  'xmlns:xlink="http://www.w3.org/1999/xlink"',
  'xmlns:dc="http://purl.org/dc/elements/1.1/"',
  'xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0"',
  'xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"',
  'xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"',
  'xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"',
].join(' ')

const XML_HEAD = '<?xml version="1.0" encoding="UTF-8"?>\n'

/** A margem de célula padrão do Word: 0,19 cm dos lados, nada em cima e embaixo. */
const CELL_PADDING =
  ' fo:padding-left="1.9mm" fo:padding-right="1.9mm" fo:padding-top="0mm" fo:padding-bottom="0mm"'

/** A borda da tabela que não declara a sua — a mesma fina da tela. */
const DEFAULT_BORDER = '0.5pt solid #000000'

/** O pacote, entrada por entrada, na ordem do arquivo: o `mimetype` primeiro e guardado. */
export function odtEntries(model: OdtModel, options: OdtExportOptions = {}): ZipEntry[] {
  const writer = new OdtWriter(model)
  const content = writer.content()
  const styles = writer.styles()
  const encoder = new TextEncoder()
  const entries: ZipEntry[] = [
    { name: 'mimetype', data: encoder.encode(MIMETYPE), stored: true },
    { name: 'content.xml', data: encoder.encode(content) },
    { name: 'styles.xml', data: encoder.encode(styles) },
    { name: 'meta.xml', data: encoder.encode(metaXml(model, options)) },
  ]
  for (const picture of writer.pictures.list()) {
    entries.push({ name: picture.path, data: picture.bytes, stored: true })
  }
  for (const formula of writer.formulas.list()) {
    entries.push({ name: `${formula.path}/content.xml`, data: encoder.encode(formula.content) })
  }
  entries.push({
    name: 'META-INF/manifest.xml',
    data: encoder.encode(manifestXml(writer.pictures.list(), writer.formulas.list())),
  })
  return entries
}

/** O arquivo `.odt` pronto. Sem `deflate`, tudo sai guardado — ainda um pacote válido. */
export function exportOdt(model: OdtModel, options: OdtExportOptions = {}, deflate?: Deflate): Uint8Array {
  return zip(odtEntries(model, options), deflate)
}

interface Picture {
  readonly path: string
  readonly mime: string
  readonly bytes: Uint8Array
  /** Largura e altura em pixels, lidas do cabeçalho do arquivo, quando dá. */
  readonly size: { readonly width: number; readonly height: number } | null
}

/** As imagens do pacote, uma vez cada — a mesma figura repetida é um arquivo só. */
class PictureBook {
  private readonly bySource = new Map<string, Picture>()

  add(src: unknown): Picture | null {
    const data = imageData(src)
    if (data === null) return null
    const known = this.bySource.get(data.base64)
    if (known !== undefined) return known
    let bytes: Uint8Array
    try {
      bytes = Uint8Array.from(atob(data.base64), (char) => char.charCodeAt(0))
    } catch {
      return null
    }
    const picture: Picture = {
      path: `Pictures/image${this.bySource.size + 1}.${imageExtension(data.mime)}`,
      mime: data.mime,
      bytes,
      size: pixelSizeOf(bytes),
    }
    this.bySource.set(data.base64, picture)
    return picture
  }

  list(): Picture[] {
    return [...this.bySource.values()]
  }
}

interface Formula {
  /** A pasta do objeto no pacote, `Object 1` — sem a barra do fim. */
  readonly path: string
  /** O `content.xml` do objeto: o MathML da equação, que é o que o ODF guarda nele. */
  readonly content: string
}

/**
 * As equações do pacote: cada uma é um objeto de fórmula embutido, uma subpasta
 * com o MathML — como o LibreOffice Math grava. Sem a imagem de substituição
 * (`ObjectReplacements/`): quem abre o arquivo desenha a fórmula.
 */
class FormulaBook {
  private readonly formulas: Formula[] = []

  add(tree: MathElement, display: boolean): Formula {
    const math = mathMlToString({ ...tree, attrs: { ...tree.attrs, display: display ? 'block' : 'inline' } })
    const formula: Formula = { path: `Object ${this.formulas.length + 1}`, content: `${XML_HEAD}${math}` }
    this.formulas.push(formula)
    return formula
  }

  list(): readonly Formula[] {
    return this.formulas
  }
}

/**
 * O tamanho do quadro da equação, estimado: quem desenha a fórmula a ajusta, mas
 * o quadro precisa de um tamanho para a linha não pular quando ninguém desenha.
 * A largura sai do texto; a altura, de quantas coisas a equação empilha.
 */
export function formulaSizeMm(tree: MathElement): { width: number; height: number } {
  const STACKED = new Set(['mfrac', 'munder', 'mover', 'munderover', 'mtable'])
  const levels = (node: MathElement): number => {
    const inner = Math.max(
      0,
      ...node.children.map((child) => (typeof child === 'string' ? 0 : levels(child))),
    )
    if (node.tag === 'mtable') return inner + Math.max(0, node.children.length - 1)
    return inner + (STACKED.has(node.tag) ? 1 : 0)
  }
  const chars = [...mathText(tree)].length
  return { width: Math.max(3, Math.min(170, chars * 2.2 + 1)), height: Math.min(120, 5 + levels(tree) * 3.5) }
}

/** A largura e a altura gravadas no PNG, no GIF ou no JPEG; `null` nos outros. */
export function pixelSizeOf(bytes: Uint8Array): { width: number; height: number } | null {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  if (bytes.length >= 24 && bytes[0] === 0x89 && bytes[1] === 0x50) {
    return { width: view.getUint32(16), height: view.getUint32(20) }
  }
  if (bytes.length >= 10 && bytes[0] === 0x47 && bytes[1] === 0x49) {
    return { width: view.getUint16(6, true), height: view.getUint16(8, true) }
  }
  if (bytes.length >= 4 && bytes[0] === 0xff && bytes[1] === 0xd8) {
    let at = 2
    while (at + 9 < bytes.length && bytes[at] === 0xff) {
      const marker = bytes[at + 1]!
      const length = view.getUint16(at + 2)
      // SOF0 a SOF15, menos DHT (C4), JPG (C8) e DAC (CC).
      if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
        return { width: view.getUint16(at + 7), height: view.getUint16(at + 5) }
      }
      at += 2 + length
    }
  }
  return null
}

/** Para onde vai a quebra ou a página mestra pendente: o próximo bloco do fluxo. */
interface Pending {
  breakBefore: 'page' | null
  master: { readonly name: string; readonly pageNumber: number | 'auto' } | null
}

class OdtWriter {
  readonly pictures = new PictureBook()
  readonly formulas = new FormulaBook()
  readonly fonts = new FontBook()
  private readonly sheet: StyleSheet
  private readonly source: ExportSource
  private readonly sections: PageSetup[]
  /** O nome ODF de cada estilo do documento, por id. */
  private readonly styleNames = new Map<string, string>()
  /** Os estilos de título que o documento não define e o arquivo precisa. */
  private readonly fallbackHeadings = new Map<number, string>()
  private frameCount = 0
  private tableCount = 0
  private sectionCount = 0

  /** A seção em que o texto está — a largura da tabela sem larguras declaradas sai dela. */
  currentPage: PageSetup

  constructor(readonly model: OdtModel) {
    this.currentPage = model.page
    this.sheet = model.styles
    this.source = prepareExport(model, { keepComments: true })
    this.sections = effectiveSections(model.page, model.sections)
    const used = new Set<string>(['Standard'])
    for (const style of Object.values(this.sheet.styles)) {
      let name = style.id.replace(/[^\p{L}\p{N}_.-]/gu, '_')
      if (!/^[\p{L}_]/u.test(name)) name = `_${name}`
      while (used.has(name)) name = `${name}_`
      used.add(name)
      this.styleNames.set(style.id, name)
    }
  }

  content(): string {
    const book = new StyleBook('')
    const pending: Pending = { breakBefore: null, master: null }
    const renderer = new Renderer(this, book, pending, this.source, true)
    const top = this.source.doc.content ?? []
    const known = this.model.sections ?? []
    const indexes = blockSections(top.map(sectionBreakInJson), known)

    const body: string[] = []
    let at = 0
    while (at < top.length) {
      const index = indexes[at]!
      const group: DocumentNode[] = []
      while (at < top.length && indexes[at] === index) group.push(top[at++]!)
      const section = this.sections[index] ?? this.model.page
      this.currentPage = section
      const first = body.length === 0
      if (first || (section.start ?? 'nextPage') !== 'continuous') {
        pending.master = {
          name: masterName(index),
          pageNumber: section.pageNumberStart ?? (first ? 1 : 'auto'),
        }
      }
      const inner = renderer.blocks(group)
      const columns = Math.round(section.columns?.count ?? 1)
      body.push(columns > 1 ? this.columnSection(book, section, inner) : inner)
    }
    // O documento sem bloco nenhum ainda precisa da página mestra no primeiro parágrafo.
    if (body.join('') === '' || pending.master !== null) body.push(renderer.emptyParagraph())

    return (
      `${XML_HEAD}<office:document-content ${NAMESPACES} office:version="1.3">` +
      this.fonts.xml() +
      `<office:automatic-styles>${book.xml()}</office:automatic-styles>` +
      `<office:body><office:text>${body.join('')}</office:text></office:body></office:document-content>`
    )
  }

  /** O `styles.xml`: os estilos nomeados, as notas, as páginas e as faixas. */
  styles(): string {
    const book = new StyleBook('M')
    const pending: Pending = { breakBefore: null, master: null }
    const renderer = new Renderer(this, book, pending, this.source, false)
    const layouts: string[] = []
    const masters: string[] = []

    this.sections.forEach((section, index) => {
      const layout = `pm${index + 1}`
      layouts.push(pageLayoutXml(layout, section))
      const band = (key: 'header' | 'footer', element: string, value: Band | null): string => {
        if (!bandHasContent(value)) return ''
        return `<style:${element}>${renderer.band(value, section, key)}</style:${element}>`
      }
      const header = section.headerBand ?? plainBand(section.header)
      const footer = section.footerBand ?? plainBand(section.footer)
      const parts = [band('header', 'header', header)]
      if (section.evenAndOddHeaders === true)
        parts.push(band('header', 'header-left', section.evenHeaderBand))
      if (section.titlePage === true) parts.push(band('header', 'header-first', section.firstHeaderBand))
      parts.push(band('footer', 'footer', footer))
      if (section.evenAndOddHeaders === true)
        parts.push(band('footer', 'footer-left', section.evenFooterBand))
      if (section.titlePage === true) parts.push(band('footer', 'footer-first', section.firstFooterBand))
      masters.push(
        `<style:master-page${attr('style:name', masterName(index))}${attr('style:page-layout-name', layout)}>` +
          `${parts.join('')}</style:master-page>`,
      )
    })

    const named = this.namedStyles()
    return (
      `${XML_HEAD}<office:document-styles ${NAMESPACES} office:version="1.3">` +
      this.fonts.xml() +
      `<office:styles>${named}${notesConfiguration(this.model)}${OUTLINE_STYLE}</office:styles>` +
      `<office:automatic-styles>${layouts.join('')}${book.xml()}</office:automatic-styles>` +
      `<office:master-styles>${masters.join('')}</office:master-styles></office:document-styles>`
    )
  }

  /** O nome do estilo de parágrafo que o bloco aponta, como a cascata o escolhe. */
  paragraphStyleName(node: DocumentNode): string {
    const id = node.attrs?.['styleId']
    if (
      typeof id === 'string' &&
      this.styleNames.has(id) &&
      this.sheet.styles[id]?.type === StyleType.Paragraph
    ) {
      return this.styleNames.get(id)!
    }
    const level = Number(node.attrs?.['level'])
    if (node.type === 'heading' && Number.isInteger(level) && level >= 1) {
      const name = `heading ${level}`
      const own = Object.values(this.sheet.styles).find(
        (style) => style.type === StyleType.Paragraph && style.name.toLowerCase() === name,
      )
      if (own !== undefined) return this.styleNames.get(own.id)!
      let fallback = this.fallbackHeadings.get(level)
      if (fallback === undefined) {
        fallback = `Heading_${level}`
        this.fallbackHeadings.set(level, fallback)
      }
      return fallback
    }
    const standard = this.sheet.defaults.paragraphStyleId
    return standard !== null && this.styleNames.has(standard) ? this.styleNames.get(standard)! : 'Standard'
  }

  characterStyleName(id: unknown): string | null {
    if (typeof id !== 'string') return null
    return this.sheet.styles[id]?.type === StyleType.Character ? (this.styleNames.get(id) ?? null) : null
  }

  /** O nível de estrutura do bloco (1 a 9), ou `null` para o corpo de texto. */
  outlineLevelOf(node: DocumentNode): number | null {
    if (node.type === 'heading') {
      const level = Number(node.attrs?.['level'])
      return Number.isInteger(level) && level >= 1 ? Math.min(level, 10) : 1
    }
    const level = blockStyleOfNode(node, this.sheet)?.paragraph.outlineLevel
    return level !== undefined && level >= 0 && level < 9 ? level + 1 : null
  }

  nextFrame(): number {
    return ++this.frameCount
  }

  nextTable(): number {
    return ++this.tableCount
  }

  /** A seção de colunas: o ODF as põe num `text:section`, que pode começar no meio da folha. */
  private columnSection(book: StyleBook, section: PageSetup, inner: string): string {
    const count = Math.max(1, Math.round(section.columns?.count ?? 1))
    const gap = Math.max(0, section.columns?.spaceMm ?? 12.7)
    const separator =
      section.columns?.separator === true
        ? '<style:column-sep style:width="0.2mm" style:color="#000000" style:height="100%" style:vertical-align="top"/>'
        : ''
    const style = book.style(
      'section',
      'Sect',
      '',
      `<style:section-properties text:dont-balance-text-columns="false"><style:columns${attr('fo:column-count', count)}${attr('fo:column-gap', mm(gap))}>${separator}</style:columns></style:section-properties>`,
    )
    const name = `Section${++this.sectionCount}`
    return `<text:section${attr('text:style-name', style)}${attr('text:name', name)}>${inner}</text:section>`
  }

  private namedStyles(): string {
    const defaults = this.sheet.defaults
    const paragraphDefaults: ParagraphProps = {
      ...paragraphPropsOfStyle(defaults.paragraph),
      widowControl: defaults.paragraph.widowControl ?? true,
    }
    const parts = [
      `<style:default-style style:family="paragraph"><style:paragraph-properties style:tab-stop-distance="12.7mm" style:writing-mode="page"${paragraphProperties(paragraphDefaults)}/>` +
        `<style:text-properties${textProperties(characterPropsOfStyle(defaults.character), this.fonts)}/></style:default-style>`,
      '<style:style style:name="Standard" style:family="paragraph" style:class="text"/>',
    ]
    for (const style of Object.values(this.sheet.styles)) parts.push(this.namedStyle(style))
    for (const [level, name] of this.fallbackHeadings) {
      const resolved = headingStyleOf(this.sheet, level)
      parts.push(
        `<style:style${attr('style:name', name)}${attr('style:display-name', `Heading ${level}`)} style:family="paragraph" style:class="text"${attr('style:default-outline-level', level)}>` +
          `<style:paragraph-properties${paragraphProperties(paragraphPropsOfStyle(resolved.paragraph))}/>` +
          `<style:text-properties${textProperties(characterPropsOfStyle(resolved.character), this.fonts)}/></style:style>`,
      )
    }
    return parts.join('')
  }

  private namedStyle(style: StyleDefinition): string {
    const name = this.styleNames.get(style.id)!
    const character = style.type === StyleType.Character
    const parent =
      style.basedOn !== undefined && this.sheet.styles[style.basedOn]?.type === style.type
        ? this.styleNames.get(style.basedOn)
        : undefined
    const outline = style.paragraph?.outlineLevel
    const head =
      `<style:style${attr('style:name', name)}${attr('style:display-name', style.name)}` +
      `${attr('style:family', character ? 'text' : 'paragraph')}${attr('style:parent-style-name', parent)}` +
      (!character && outline !== undefined && outline >= 0 && outline < 9
        ? attr('style:default-outline-level', outline + 1)
        : '') +
      (!character && style.next !== undefined && this.styleNames.has(style.next)
        ? attr('style:next-style-name', this.styleNames.get(style.next))
        : '') +
      '>'
    const paragraph = character ? '' : paragraphProperties(paragraphPropsOfStyle(style.paragraph))
    const text = textProperties(characterPropsOfStyle(style.character), this.fonts)
    return (
      head +
      (paragraph === '' ? '' : `<style:paragraph-properties${paragraph}/>`) +
      (text === '' ? '' : `<style:text-properties${text}/>`) +
      '</style:style>'
    )
  }
}

/** A faixa tem o que desenhar — o texto, a grade, o filete ou só objetos ancorados. */
function bandHasContent(band: Band | null | undefined): band is Band {
  if (band === null || band === undefined) return false
  return band.floats.length > 0 || hasBandContent(band)
}

function masterName(index: number): string {
  return index === 0 ? 'Standard' : `Section${index + 1}`
}

/** A marca de seção que o bloco de primeiro nível carrega — a dele, ou a de um parágrafo de dentro. */
function sectionBreakInJson(node: DocumentNode): string | null {
  const own = sectionBreakOf(node)
  if (own !== null || node.type === 'paragraph' || node.type === 'heading') return own
  for (const child of node.content ?? []) {
    const found = sectionBreakInJson(child)
    if (found !== null) return found
  }
  return null
}

const NUMBER_FORMATS: Readonly<Record<string, string>> = {
  decimal: '1',
  decimalZero: '1',
  lowerRoman: 'i',
  upperRoman: 'I',
  lowerLetter: 'a',
  upperLetter: 'A',
  none: '',
}

function numberFormat(fmt: string | undefined): string {
  return NUMBER_FORMATS[fmt ?? 'decimal'] ?? '1'
}

/**
 * A folha da seção. O ODF mede a margem até o cabeçalho, e o Word até o texto:
 * com cabeçalho, a margem de cima é a distância dele, e a altura mínima dele
 * cobre o resto — o cabeçalho mais alto empurra o texto, como no Word.
 */
function pageLayoutXml(name: string, page: PageSetup): string {
  const { width, height } = pageDimensionsMm(page)
  const hasHeader =
    bandHasContent(page.headerBand ?? plainBand(page.header)) ||
    (page.titlePage === true && bandHasContent(page.firstHeaderBand)) ||
    (page.evenAndOddHeaders === true && bandHasContent(page.evenHeaderBand))
  const hasFooter =
    bandHasContent(page.footerBand ?? plainBand(page.footer)) ||
    (page.titlePage === true && bandHasContent(page.firstFooterBand)) ||
    (page.evenAndOddHeaders === true && bandHasContent(page.evenFooterBand))
  const top = hasHeader ? Math.min(page.headerDistanceMm, page.margins.top) : page.margins.top
  const bottom = hasFooter ? Math.min(page.footerDistanceMm, page.margins.bottom) : page.margins.bottom
  const properties =
    attr('fo:page-width', mm(width)) +
    attr('fo:page-height', mm(height)) +
    attr('style:print-orientation', width > height ? 'landscape' : 'portrait') +
    attr('fo:margin-top', mm(top)) +
    attr('fo:margin-bottom', mm(bottom)) +
    attr('fo:margin-left', mm(page.margins.left)) +
    attr('fo:margin-right', mm(page.margins.right)) +
    attr('style:num-format', numberFormat(page.pageNumberFormat) || '1') +
    ' style:writing-mode="lr-tb"'
  const header = hasHeader
    ? `<style:header-style><style:header-footer-properties${attr('fo:min-height', mm(Math.max(0, page.margins.top - top)))} fo:margin-bottom="0mm" style:dynamic-spacing="false"/></style:header-style>`
    : '<style:header-style/>'
  const footer = hasFooter
    ? `<style:footer-style><style:header-footer-properties${attr('fo:min-height', mm(Math.max(0, page.margins.bottom - bottom)))} fo:margin-top="0mm" style:dynamic-spacing="false"/></style:footer-style>`
    : '<style:footer-style/>'
  return `<style:page-layout${attr('style:name', name)}><style:page-layout-properties${properties}/>${header}${footer}</style:page-layout>`
}

/** Como as notas se numeram: o formato e o início do documento, e onde ficam. */
function notesConfiguration(model: OdtModel): string {
  const one = (kind: 'footnote' | 'endnote'): string => {
    const numbering = kind === 'endnote' ? model.notes?.endnotePr : model.notes?.footnotePr
    const format = numbering?.numFmt ?? (kind === 'endnote' ? 'lowerRoman' : 'decimal')
    const restart =
      numbering?.restart === 'eachPage' && kind === 'footnote'
        ? 'page'
        : numbering?.restart === 'eachSect'
          ? 'chapter'
          : 'document'
    // O LibreOffice conta o início a partir de zero: 0 é "começa em 1".
    return (
      `<text:notes-configuration${attr('text:note-class', kind)}${attr('style:num-format', numberFormat(format) || '1')}` +
      `${attr('text:start-value', Math.max(0, (numbering?.start ?? 1) - 1))}` +
      `${kind === 'footnote' ? ' text:footnotes-position="page"' : ''}${attr('text:start-numbering-at', restart)}/>`
    )
  }
  return one('footnote') + one('endnote')
}

/** Os títulos sem número: o estilo de estrutura padrão do ODF numera, e o documento não pediu. */
const OUTLINE_STYLE = `<text:outline-style style:name="Outline">${Array.from(
  { length: 10 },
  (_, index) =>
    `<text:outline-level-style text:level="${index + 1}" style:num-format=""><style:list-level-properties text:list-level-position-and-space-mode="label-alignment"><style:list-level-label-alignment text:label-followed-by="nothing"/></style:list-level-properties></text:outline-level-style>`,
).join('')}</text:outline-style>`

class Renderer {
  /** Dentro de célula, nota, caixa de texto ou faixa a quebra pendente espera. */
  private nested = 0
  /** Dentro de nota, faixa ou caixa de texto, onde não pode haver nota. */
  private noteless = 0
  /** Renderizando os objetos de uma faixa, onde `{n}` é o número da página. */
  private inBand = false
  /** O que uma lista não pode conter (tabela) e sai logo depois dela. */
  private hoisted: string[] = []
  private readonly bookmarkEnds = new Set<string>()
  private readonly bookmarkNames = new Map<string, string>()
  private readonly usedBookmarks = new Set<string>()
  private readonly commentEnds = new Set<string>()
  private readonly openComments = new Set<string>()

  constructor(
    private readonly writer: OdtWriter,
    private readonly book: StyleBook,
    private readonly pending: Pending,
    private readonly source: ExportSource,
    body: boolean,
  ) {
    if (!body) {
      this.nested = 1
      this.noteless = 1
    }
    const scan = (node: DocumentNode): void => {
      if (node.type === 'bookmarkEnd') this.bookmarkEnds.add(String(node.attrs?.['bid'] ?? ''))
      if (node.type === 'commentEnd') this.commentEnds.add(String(node.attrs?.['cid'] ?? ''))
      for (const child of node.content ?? []) scan(child)
    }
    if (body) scan(source.doc)
  }

  blocks(nodes: readonly DocumentNode[]): string {
    return nodes.map((node) => this.block(node)).join('')
  }

  block(node: DocumentNode): string {
    switch (node.type) {
      case 'paragraph':
      case 'heading':
        return isSectionMarkOnly(node) ? '' : this.paragraph(node)
      case 'bulletList':
      case 'orderedList': {
        const outer = this.hoisted
        this.hoisted = []
        const list = this.list(node, 1)
        const after = this.hoisted.join('')
        this.hoisted = outer
        return list + after
      }
      case 'table':
        return this.table(node)
      case 'codeBlock':
        return this.codeBlock(node)
      case 'horizontalRule':
        return `<text:p${attr('text:style-name', this.paragraphStyle({ type: 'paragraph' }, { borderBottom: DEFAULT_BORDER }))}/>`
      case 'pageBreak':
        this.pending.breakBefore = 'page'
        return ''
      default:
        // O sumário, a citação e o bloco que esta exportação não conhece: o conteúdo.
        return node.content === undefined ? '' : this.blocks(node.content)
    }
  }

  /** Um parágrafo vazio, que leva o que estiver pendente. */
  emptyParagraph(): string {
    return `<text:p${attr('text:style-name', this.paragraphStyle({ type: 'paragraph' }, {}))}/>`
  }

  private paragraph(node: DocumentNode): string {
    const attrs = node.attrs ?? {}
    const breakAfter =
      attrs['breakAfter'] === true ? 'page' : attrs['columnBreakAfter'] === true ? 'column' : null
    const style = this.paragraphStyle(node, breakAfter === null ? {} : { breakAfter })
    const frames = floatsOf(attrs)
      .map((object) => this.floating(object))
      .join('')
    const content = frames + this.inline(node.content ?? [])
    const level = this.writer.outlineLevelOf(node)
    return level === null
      ? `<text:p${attr('text:style-name', style)}>${content}</text:p>`
      : `<text:h${attr('text:style-name', style)}${attr('text:outline-level', level)}>${content}</text:h>`
  }

  /** O estilo automático do bloco: o nomeado dele por baixo, a formatação direta por cima. */
  private paragraphStyle(node: DocumentNode, extra: ParagraphProps): string {
    const attrs = node.attrs ?? {}
    const parent = this.writer.paragraphStyleName(node)
    const indentLevel = finite(attrs['indent']) ?? 0
    const indent = finite(attrs['indentMm'])
    const hanging = finite(attrs['hangingMm'])
    const props: ParagraphProps = {
      align: attrs['textAlign'],
      marginLeftMm:
        indent !== null || indentLevel > 0 ? (indent ?? 0) + Math.max(0, indentLevel) * INDENT_STEP_MM : null,
      marginRightMm: finite(attrs['indentRightMm']),
      textIndentMm: hanging !== null && hanging > 0 ? -hanging : finite(attrs['firstLineMm']),
      spaceBeforePt: finite(attrs['spaceBefore']),
      spaceAfterPt: finite(attrs['spaceAfter']),
      ...lineSpacingOfAttr(attrs['lineHeight'], attrs['fontFamily']),
      keepNext: typeof attrs['keepNext'] === 'boolean' ? attrs['keepNext'] : undefined,
      keepLines: typeof attrs['keepLines'] === 'boolean' ? attrs['keepLines'] : undefined,
      widowControl: typeof attrs['widowControl'] === 'boolean' ? attrs['widowControl'] : undefined,
      background: attrs['background'],
      ...extra,
    }
    let master = ''
    if (this.nested === 0) {
      if (this.pending.breakBefore !== null) props.breakBefore = this.pending.breakBefore
      if (this.pending.master !== null) {
        master = attr('style:master-page-name', this.pending.master.name)
        props.pageNumber = this.pending.master.pageNumber
        props.breakBefore = null
      }
      this.pending.breakBefore = null
      this.pending.master = null
    }
    const paragraph = paragraphProperties(props)
    const text = textProperties(
      { fontFamily: attrs['fontFamily'], fontSize: attrs['fontSize'] },
      this.writer.fonts,
    )
    if (paragraph === '' && text === '' && master === '') return parent
    return this.book.style(
      'paragraph',
      'P',
      `${attr('style:parent-style-name', parent)}${master}`,
      (paragraph === '' ? '' : `<style:paragraph-properties${paragraph}/>`) +
        (text === '' ? '' : `<style:text-properties${text}/>`),
    )
  }

  private codeBlock(node: DocumentNode): string {
    const style = this.paragraphStyle({ type: 'paragraph' }, { spaceBeforePt: 0, spaceAfterPt: 0 })
    const mono = this.book.style(
      'text',
      'T',
      '',
      `<style:text-properties${textProperties({ mono: true }, this.writer.fonts)}/>`,
    )
    return plainText(node)
      .split('\n')
      .map(
        (line) =>
          `<text:p${attr('text:style-name', style)}><text:span${attr('text:style-name', mono)}>${odfText(line)}</text:span></text:p>`,
      )
      .join('')
  }

  private list(node: DocumentNode, depth: number): string {
    const styleName = depth === 1 ? this.listStyle(node) : null
    const ordered = node.type === 'orderedList'
    const items = (node.content ?? []).filter((child) => child.type === 'listItem')
    const body = items
      .map((item, index) => {
        const value = this.source.itemOf.get(item)?.value
        // Todo item que abre lista numerada diz o número dele: o Word continua a
        // contagem entre listas separadas, e o ODF recomeçaria do 1.
        const start = ordered && index === 0 && value !== undefined ? attr('text:start-value', value) : ''
        const inner = (item.content ?? [])
          .map((child) => {
            if (child.type === 'paragraph' || child.type === 'heading') {
              return isSectionMarkOnly(child) ? '' : this.paragraph(child)
            }
            if (child.type === 'bulletList' || child.type === 'orderedList')
              return this.list(child, depth + 1)
            // O ODF não deixa tabela dentro de item de lista: ela sai depois da lista.
            this.hoisted.push(this.block(child))
            return ''
          })
          .join('')
        return `<text:list-item${start}>${inner === '' ? '<text:p/>' : inner}</text:list-item>`
      })
      .join('')
    return `<text:list${attr('text:style-name', styleName)}>${body}</text:list>`
  }

  /**
   * O estilo da lista de fora, com um nível por profundidade: cada nível diz o
   * que a conta (`numberLists`) decidiu para a primeira lista daquela
   * profundidade — formato, texto do número, recuo.
   */
  private listStyle(node: DocumentNode): string {
    const byDepth = new Map<number, ListInfo>()
    const visit = (list: DocumentNode, depth: number): void => {
      const info = this.source.listOf.get(list)
      if (info !== undefined && !byDepth.has(depth)) byDepth.set(depth, info)
      for (const item of list.content ?? []) {
        for (const child of item.content ?? []) {
          if (child.type === 'bulletList' || child.type === 'orderedList') visit(child, depth + 1)
        }
      }
    }
    visit(node, 1)
    const levels = [...byDepth.entries()]
      .filter(([depth]) => depth <= 10)
      .sort(([a], [b]) => a - b)
      .map(([depth, info]) => listLevelXml(depth, info))
      .join('')
    return this.book.list(levels)
  }

  private table(node: DocumentNode): string {
    const rows = (node.content ?? []).filter((row) => row.type === 'tableRow')
    const grid = tableGrid(rows)
    if (grid.columns === 0) return ''
    const available = contentWidthMm(this.writer.currentPage)
    const declared = columnWidthsMm(rows, grid.columns)
    const widths = declared ?? Array.from({ length: grid.columns }, () => available / grid.columns)
    const total = widths.reduce((sum, width) => sum + width, 0)

    let master = ''
    const tableProps: string[] = [attr('style:width', mm(total)), ' table:align="left"']
    if (this.nested === 0) {
      if (this.pending.master !== null) {
        master = attr('style:master-page-name', this.pending.master.name)
        tableProps.push(attr('style:page-number', this.pending.master.pageNumber))
      } else if (this.pending.breakBefore !== null) {
        tableProps.push(' fo:break-before="page"')
      }
      this.pending.master = null
      this.pending.breakBefore = null
    }
    const tableStyle = this.book.style(
      'table',
      'Ta',
      master,
      `<style:table-properties${tableProps.join('')}/>`,
    )
    const columns = widths
      .map((width) => {
        const style = this.book.style(
          'table-column',
          'co',
          '',
          `<style:table-column-properties${attr('style:column-width', mm(width))}/>`,
        )
        return `<table:table-column${attr('table:style-name', style)}/>`
      })
      .join('')

    this.nested++
    const body = grid.rows
      .map((slots) => {
        const cells = slots
          .map((slot) =>
            slot === null ? '<table:covered-table-cell/>' : this.cell(slot.node, slot.colspan, slot.rowspan),
          )
          .join('')
        return `<table:table-row>${cells}</table:table-row>`
      })
      .join('')
    this.nested--

    const name = `Table${this.writer.nextTable()}`
    return `<table:table${attr('table:name', name)}${attr('table:style-name', tableStyle)}>${columns}${body}</table:table>`
  }

  private cell(node: DocumentNode, colspan: number, rowspan: number): string {
    const attrs = node.attrs ?? {}
    const borders = cellBordersFromAttr(attrs['borders'])
    const side = (border: CellBorder | null): string =>
      border === null
        ? DEFAULT_BORDER
        : border.style === 'none'
          ? 'none'
          : `${Math.round(border.widthPt * 100) / 100}pt ${border.style === 'double' ? 'double' : border.style === 'dashed' ? 'dashed' : border.style === 'dotted' ? 'dotted' : 'solid'} ${border.color}`
    const shading = odfColor(attrs['shading'])
    const style = this.book.style(
      'table-cell',
      'ce',
      '',
      `<style:table-cell-properties${CELL_PADDING}` +
        `${attr('fo:border-top', side(borders.top))}${attr('fo:border-bottom', side(borders.bottom))}` +
        `${attr('fo:border-left', side(borders.left))}${attr('fo:border-right', side(borders.right))}` +
        `${attr('fo:background-color', shading)}/>`,
    )
    const content = this.blocks(node.content ?? [])
    return (
      `<table:table-cell${attr('table:style-name', style)} office:value-type="string"` +
      `${colspan > 1 ? attr('table:number-columns-spanned', colspan) : ''}` +
      `${rowspan > 1 ? attr('table:number-rows-spanned', rowspan) : ''}>` +
      `${content === '' ? '<text:p/>' : content}</table:table-cell>` +
      '<table:covered-table-cell/>'.repeat(colspan - 1)
    )
  }

  inline(nodes: readonly DocumentNode[]): string {
    return nodes.map((node) => this.inlineNode(node)).join('')
  }

  private inlineNode(node: DocumentNode): string {
    switch (node.type) {
      case 'text':
        // Na faixa (e nas caixas dela), `{n}` e `{total}` são os campos de página.
        return this.wrap(
          this.nested > 0 && this.inBand ? bandText(node.text ?? '') : odfText(node.text ?? ''),
          node.marks ?? [],
        )
      case 'hardBreak':
        return '<text:line-break/>'
      case 'image': {
        const frame = this.image(node)
        return frame === ''
          ? ''
          : this.wrap(
              frame,
              (node.marks ?? []).filter((mark) => mark.type === 'link'),
            )
      }
      case 'noteRef':
        return this.note(node)
      case 'field':
        return this.wrap(this.field(node), node.marks ?? [])
      case 'math':
        return this.wrap(this.formula(node), node.marks ?? [])
      case 'bookmarkStart':
        return this.bookmarkStart(node)
      case 'bookmarkEnd': {
        const name = this.bookmarkNames.get(String(node.attrs?.['bid'] ?? ''))
        return name === undefined ? '' : `<text:bookmark-end${attr('text:name', name)}/>`
      }
      case 'commentStart':
        return this.commentStart(node)
      case 'commentEnd': {
        const id = String(node.attrs?.['cid'] ?? '')
        return this.openComments.has(id)
          ? `<office:annotation-end${attr('office:name', annotationName(id))}/>`
          : ''
      }
      default:
        return node.content === undefined
          ? this.wrap(odfText(node.text ?? ''), node.marks ?? [])
          : this.inline(node.content)
    }
  }

  /** O trecho com as marcas dele: um `text:span` com estilo automático, e o link por fora. */
  private wrap(inner: string, marks: NonNullable<DocumentNode['marks']>): string {
    if (inner === '') return ''
    const props: CharacterProps = {}
    let characterStyle: string | null = null
    let href: string | null = null
    for (const mark of marks) {
      const attrs = mark.attrs ?? {}
      const on = attrs['off'] !== true
      switch (mark.type) {
        case 'bold':
          props.bold = on
          break
        case 'italic':
          props.italic = on
          break
        case 'underline':
          props.underline = on
          break
        case 'strike':
          props.strike = on
          break
        case 'code':
          props.mono = true
          break
        case 'superscript':
          props.position = 'super'
          break
        case 'subscript':
          props.position = 'sub'
          break
        case 'caps':
          props.caps = true
          break
        case 'smallCaps':
          props.smallCaps = true
          break
        case 'highlight':
          props.background = odfColor(attrs['color']) ?? '#ffff00'
          break
        case 'charStyle':
          characterStyle = this.writer.characterStyleName(attrs['styleId'])
          break
        case 'textStyle':
          if (attrs['color'] !== undefined) props.color = attrs['color']
          if (attrs['backgroundColor'] !== undefined) props.background = attrs['backgroundColor']
          if (attrs['fontFamily'] !== undefined) props.fontFamily = attrs['fontFamily']
          if (attrs['fontSize'] !== undefined) props.fontSize = attrs['fontSize']
          break
        case 'link':
          href = safeHref(mark)
          break
        default:
          break
      }
    }
    let out = inner
    const text = textProperties(props, this.writer.fonts)
    if (text !== '') {
      const style = this.book.style(
        'text',
        'T',
        attr('style:parent-style-name', characterStyle),
        `<style:text-properties${text}/>`,
      )
      out = `<text:span${attr('text:style-name', style)}>${out}</text:span>`
    } else if (characterStyle !== null) {
      out = `<text:span${attr('text:style-name', characterStyle)}>${out}</text:span>`
    }
    if (href !== null) out = `<text:a xlink:type="simple"${attr('xlink:href', href)}>${out}</text:a>`
    return out
  }

  private field(node: DocumentNode): string {
    const instr = String(node.attrs?.['instr'] ?? '')
    const result = odfText(typeof node.attrs?.['result'] === 'string' ? node.attrs['result'] : '')
    switch (fieldKind(instr)) {
      case 'PAGE':
        return `<text:page-number text:select-page="current">${result}</text:page-number>`
      case 'NUMPAGES':
        return `<text:page-count>${result}</text:page-count>`
      case 'DATE':
        return `<text:date>${result}</text:date>`
      case 'TIME':
        return `<text:time>${result}</text:time>`
      default:
        return result
    }
  }

  private bookmarkStart(node: DocumentNode): string {
    const name = typeof node.attrs?.['name'] === 'string' ? node.attrs['name'] : ''
    // Dois marcadores com o mesmo nome não abrem no LibreOffice: o segundo sai.
    if (name === '' || this.usedBookmarks.has(name)) return ''
    this.usedBookmarks.add(name)
    const bid = String(node.attrs?.['bid'] ?? '')
    if (!this.bookmarkEnds.has(bid)) return `<text:bookmark${attr('text:name', name)}/>`
    this.bookmarkNames.set(bid, name)
    return `<text:bookmark-start${attr('text:name', name)}/>`
  }

  /** O comentário e as respostas dele, como anotações; a primeira cobre o trecho até a ponta final. */
  private commentStart(node: DocumentNode): string {
    const id = String(node.attrs?.['cid'] ?? '')
    const library = this.writer.model.comments ?? []
    const root = library.find((comment) => comment.id === id)
    if (root === undefined || this.openComments.has(id)) return ''
    const ranged = this.commentEnds.has(id)
    if (ranged) this.openComments.add(id)
    const annotation = (comment: (typeof library)[number], name: string | null): string => {
      const date = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/.test(comment.date) ? comment.date : null
      const paragraphs = comment.paragraphs.length === 0 ? [''] : comment.paragraphs
      return (
        `<office:annotation${attr('office:name', name)}${root.done ? ' loext:resolved="true"' : ''}>` +
        `<dc:creator>${xml(comment.author)}</dc:creator>` +
        (date === null ? '' : `<dc:date>${xml(date)}</dc:date>`) +
        (comment.initials === undefined || comment.initials === ''
          ? ''
          : `<meta:creator-initials>${xml(comment.initials)}</meta:creator-initials>`) +
        paragraphs.map((text) => `<text:p>${odfText(text)}</text:p>`).join('') +
        '</office:annotation>'
      )
    }
    const replies = library.filter((comment) => comment.parentId === id)
    return (
      annotation(root, ranged ? annotationName(id) : null) +
      replies.map((reply) => annotation(reply, null)).join('')
    )
  }

  private note(node: DocumentNode): string {
    const note = this.source.noteOf.get(node)
    // Nota não cabe em nota, nem em faixa, nem em caixa de texto: o ODF não deixa.
    if (note === undefined || this.noteless > 0) return ''
    const endnote = note.kind === NoteKind.Endnote
    const custom = typeof node.attrs?.['mark'] === 'string' && node.attrs['mark'] !== ''
    this.nested++
    this.noteless++
    const body = this.blocks(note.body)
    this.noteless--
    this.nested--
    return (
      `<text:note${attr('text:id', note.id)}${attr('text:note-class', endnote ? 'endnote' : 'footnote')}>` +
      `<text:note-citation${custom ? attr('text:label', note.label) : ''}>${xml(note.label)}</text:note-citation>` +
      `<text:note-body>${body === '' ? '<text:p/>' : body}</text:note-body></text:note>`
    )
  }

  private image(node: DocumentNode): string {
    const picture = this.writer.pictures.add(node.attrs?.['src'])
    if (picture === null) return ''
    const width = finite(node.attrs?.['width'])
    const height = finite(node.attrs?.['height'])
    const size = sizeMm(width, height, picture)
    const style = this.book.style(
      'graphic',
      'fr',
      '',
      '<style:graphic-properties style:vertical-pos="top" style:vertical-rel="baseline" style:horizontal-pos="center" style:horizontal-rel="paragraph" fo:padding="0mm" fo:border="none" style:mirror="none"/>',
    )
    const alt = typeof node.attrs?.['alt'] === 'string' ? node.attrs['alt'] : ''
    const number = this.writer.nextFrame()
    return (
      `<draw:frame${attr('draw:style-name', style)}${attr('draw:name', `Image${number}`)} text:anchor-type="as-char"` +
      `${attr('svg:width', mm(size.width))}${attr('svg:height', mm(size.height))}${attr('draw:z-index', number)}>` +
      `<draw:image${attr('xlink:href', picture.path)} xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>` +
      `${alt === '' ? '' : `<svg:desc>${xml(alt)}</svg:desc>`}</draw:frame>`
    )
  }

  /** A equação: um objeto de fórmula no texto, como um caractere. */
  private formula(node: DocumentNode): string {
    const tree = sanitizeMathMl(typeof node.attrs?.['mathml'] === 'string' ? node.attrs['mathml'] : '')
    if (tree === null) return ''
    const formula = this.writer.formulas.add(tree, node.attrs?.['display'] === true)
    const size = formulaSizeMm(tree)
    const style = this.book.style(
      'graphic',
      'fr',
      '',
      '<style:graphic-properties style:vertical-pos="middle" style:vertical-rel="text" fo:padding="0mm" fo:border="none"/>',
    )
    const latex = latexOfEquation(node.attrs)
    const number = this.writer.nextFrame()
    return (
      `<draw:frame${attr('draw:style-name', style)}${attr('draw:name', `Equation${number}`)} text:anchor-type="as-char"` +
      `${attr('svg:width', mm(size.width))}${attr('svg:height', mm(size.height))}${attr('draw:z-index', number)}>` +
      `<draw:object${attr('xlink:href', `./${formula.path}`)} xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>` +
      `${latex === '' ? '' : `<svg:desc>${xml(latex)}</svg:desc>`}</draw:frame>`
    )
  }

  /** O objeto flutuante, ancorado ao parágrafo na posição que o arquivo dá — o melhor que dá. */
  private floating(object: FloatingObject): string {
    if (object.kind === 'rule') return ''
    const picture = object.kind === 'image' ? this.writer.pictures.add(object.src) : null
    if (object.kind === 'image' && picture === null) return ''
    const horizontalRel = RELATIVE[object.hFrom] ?? 'paragraph'
    const verticalRel = RELATIVE[object.vFrom] ?? 'paragraph'
    const hAlign = ['left', 'center', 'right', 'inside', 'outside'].includes(object.hAlign ?? '')
      ? object.hAlign!
      : null
    const vAlign =
      object.vAlign === 'center'
        ? 'middle'
        : ['top', 'bottom'].includes(object.vAlign ?? '')
          ? object.vAlign!
          : null
    const wrap = object.behind ? 'run-through' : (WRAP[object.wrap] ?? 'parallel')
    const style = this.book.style(
      'graphic',
      'fr',
      '',
      `<style:graphic-properties${attr('style:wrap', wrap)}${object.behind ? ' style:run-through="background"' : ' style:run-through="foreground"'}` +
        `${attr('style:horizontal-pos', hAlign ?? 'from-left')}${attr('style:horizontal-rel', horizontalRel)}` +
        `${attr('style:vertical-pos', vAlign ?? 'from-top')}${attr('style:vertical-rel', verticalRel)}` +
        // Sem preenchimento declarado é transparente: o padrão do LibreOffice pinta de azul.
        (odfColor(object.fill) === null
          ? ' fo:background-color="transparent" draw:fill="none"'
          : `${attr('fo:background-color', odfColor(object.fill))} draw:fill="solid"${attr('draw:fill-color', odfColor(object.fill))}`) +
        `${attr('fo:border', object.line !== undefined && odfColor(object.line) !== null ? `${Math.max(0.25, object.lineWidthPt ?? 0.75)}pt solid ${odfColor(object.line)}` : 'none')}` +
        ' fo:padding="0mm"/>',
    )
    const x = (object.hOffsetMm ?? 0) + (object.dxMm ?? 0)
    const y = (object.vOffsetMm ?? 0) + (object.dyMm ?? 0)
    const number = this.writer.nextFrame()
    const position = `${hAlign === null ? attr('svg:x', mm(x)) : ''}${vAlign === null ? attr('svg:y', mm(y)) : ''}`
    const width = attr('svg:width', mm(Math.max(1, object.widthMm)))
    const height = mm(Math.max(1, object.heightMm))
    let size = `${width}${attr('svg:height', height)}`
    let inner: string
    if (picture !== null) {
      inner = `<draw:image${attr('xlink:href', picture.path)} xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/>`
    } else {
      this.nested++
      this.noteless++
      const blocks = this.blocks(object.content ?? [])
      this.noteless--
      this.nested--
      // A caixa cresce com o texto, como no Word: a altura dada é o mínimo, e a
      // fonte substituta um pouco maior não esconde o que não coube.
      size = width
      inner = `<draw:text-box${attr('fo:min-height', height)}>${blocks === '' ? '<text:p/>' : blocks}</draw:text-box>`
    }
    return (
      `<draw:frame${attr('draw:style-name', style)}${attr('draw:name', `Frame${number}`)} text:anchor-type="paragraph"` +
      `${position}${size}${attr('draw:z-index', number)}>${inner}</draw:frame>`
    )
  }

  /**
   * O cabeçalho ou rodapé. A grade vira tabela; as três colunas, um parágrafo
   * com tabulações de centro e de direita quando cada uma tem uma linha só, ou
   * parágrafos alinhados quando têm mais.
   */
  band(band: Band, page: PageSetup, kind: 'header' | 'footer'): string {
    const width = contentWidthMm(page)
    const parts: string[] = []
    if (band.rows.length > 0) parts.push(this.bandGrid(band, width))
    const columns = [
      { pieces: band.left, align: 'left' },
      { pieces: band.center, align: 'center' },
      { pieces: band.right, align: 'right' },
    ].filter((column) => column.pieces.length > 0)
    const rule = band.rule && kind === 'header' ? { borderBottom: DEFAULT_BORDER } : {}
    if (columns.length > 0 && columns.every((column) => linesOf(column.pieces).length === 1)) {
      const stops =
        `<style:tab-stops><style:tab-stop style:type="center"${attr('style:position', mm(width / 2))}/>` +
        `<style:tab-stop style:type="right"${attr('style:position', mm(width))}/></style:tab-stops>`
      const style = this.book.style(
        'paragraph',
        'P',
        ' style:parent-style-name="Standard"',
        `<style:paragraph-properties fo:margin-top="0pt" fo:margin-bottom="0pt"${paragraphProperties(rule)}>${stops}</style:paragraph-properties>`,
      )
      const content = ['left', 'center', 'right']
        .map((align) => columns.find((column) => column.align === align))
        .map((column) => (column === undefined ? '' : this.pieces(column.pieces)))
      // Sem as tabulações do fim que não levam nada.
      while (content.length > 1 && content[content.length - 1] === '') content.pop()
      parts.push(`<text:p${attr('text:style-name', style)}>${content.join('<text:tab/>')}</text:p>`)
    } else {
      columns.forEach((column, index) => {
        const lines = linesOf(column.pieces)
        lines.forEach((line, lineIndex) => {
          const last = index === columns.length - 1 && lineIndex === lines.length - 1
          const style = this.book.style(
            'paragraph',
            'P',
            ' style:parent-style-name="Standard"',
            `<style:paragraph-properties fo:margin-top="0pt" fo:margin-bottom="0pt"${paragraphProperties({ align: column.align, ...(last ? rule : {}) })}/>`,
          )
          parts.push(`<text:p${attr('text:style-name', style)}>${this.pieces(line)}</text:p>`)
        })
      })
    }
    // Os objetos ancorados da faixa (a marca lateral, o número de página numa
    // caixa) vão no primeiro parágrafo dela, fora da grade, com a posição na folha.
    this.inBand = true
    const frames = band.floats.map((object) => this.floating(object)).join('')
    this.inBand = false
    if (frames !== '') {
      const at = parts.findIndex((part) => part.startsWith('<text:p'))
      if (at >= 0)
        parts[at] = parts[at]!.replace(/^<text:p([^>]*?)(\/?)>/, (_, attrs: string, empty: string) =>
          empty === '/' ? `<text:p${attrs}>${frames}</text:p>` : `<text:p${attrs}>${frames}`,
        )
      else {
        const style = this.book.style(
          'paragraph',
          'P',
          ' style:parent-style-name="Standard"',
          '<style:paragraph-properties fo:margin-top="0pt" fo:margin-bottom="0pt" fo:line-height="100%"/><style:text-properties fo:font-size="1pt"/>',
        )
        parts.push(`<text:p${attr('text:style-name', style)}>${frames}</text:p>`)
      }
    }
    return parts.length === 0 ? '<text:p/>' : parts.join('')
  }

  private bandGrid(band: Band, width: number): string {
    const grid = tableGrid(
      band.rows.map((row) => ({
        type: 'tableRow',
        content: row.cells.map((cell) => ({
          type: 'tableCell',
          attrs: { colspan: cell.span, rowspan: cell.rowSpan, cell },
        })),
      })),
    )
    const widths = gridWidths(grid, width)
    const tableStyle = this.book.style(
      'table',
      'Ta',
      '',
      `<style:table-properties${attr('style:width', mm(widths.reduce((a, b) => a + b, 0)))} table:align="left"/>`,
    )
    const cols = widths
      .map(
        (value) =>
          `<table:table-column${attr('table:style-name', this.book.style('table-column', 'co', '', `<style:table-column-properties${attr('style:column-width', mm(value))}/>`))}/>`,
      )
      .join('')
    const rows = grid.rows
      .map((slots) => {
        const cells = slots
          .map((slot) => {
            if (slot === null) return '<table:covered-table-cell/>'
            const cell = slot.node.attrs?.['cell'] as Band['rows'][number]['cells'][number]
            const side = (letter: string): string => (cell.borders.includes(letter) ? DEFAULT_BORDER : 'none')
            const style = this.book.style(
              'table-cell',
              'ce',
              '',
              `<style:table-cell-properties${CELL_PADDING}${attr('fo:border-top', side('t'))}${attr('fo:border-bottom', side('b'))}${attr('fo:border-left', side('l'))}${attr('fo:border-right', side('r'))}/>`,
            )
            const paragraph = this.book.style(
              'paragraph',
              'P',
              ' style:parent-style-name="Standard"',
              `<style:paragraph-properties fo:margin-top="0pt" fo:margin-bottom="0pt"${paragraphProperties({ align: cell.align })}/>`,
            )
            const lines = linesOf(cell.pieces)
            const content = (lines.length === 0 ? [[]] : lines)
              .map((line) => `<text:p${attr('text:style-name', paragraph)}>${this.pieces(line)}</text:p>`)
              .join('')
            return (
              `<table:table-cell${attr('table:style-name', style)} office:value-type="string"` +
              `${slot.colspan > 1 ? attr('table:number-columns-spanned', slot.colspan) : ''}` +
              `${slot.rowspan > 1 ? attr('table:number-rows-spanned', slot.rowspan) : ''}>${content}</table:table-cell>` +
              '<table:covered-table-cell/>'.repeat(slot.colspan - 1)
            )
          })
          .join('')
        return `<table:table-row>${cells}</table:table-row>`
      })
      .join('')
    return `<table:table${attr('table:name', `Table${this.writer.nextTable()}`)}${attr('table:style-name', tableStyle)}>${cols}${rows}</table:table>`
  }

  private pieces(pieces: readonly BandPiece[]): string {
    return pieces
      .map((piece) => {
        let inner: string
        switch (piece.kind) {
          case 'pageNumber':
            inner = '<text:page-number text:select-page="current">1</text:page-number>'
            break
          case 'totalPages':
            inner = '<text:page-count>1</text:page-count>'
            break
          case 'image':
            return this.image({
              type: 'image',
              attrs: { src: piece.src, width: piece.width, height: piece.height },
            })
          default:
            inner = piece.literal === true ? odfText(piece.text ?? '') : bandText(piece.text ?? '')
        }
        const text = textProperties(
          {
            ...(piece.bold ? { bold: true } : {}),
            ...(piece.italic ? { italic: true } : {}),
            color: piece.color,
            fontSize: piece.fontSize,
            fontFamily: piece.fontFamily,
          },
          this.writer.fonts,
        )
        if (text === '' || inner === '') return inner
        const style = this.book.style('text', 'T', '', `<style:text-properties${text}/>`)
        return `<text:span${attr('text:style-name', style)}>${inner}</text:span>`
      })
      .join('')
  }
}

/**
 * As larguras das colunas da grade do cabeçalho, em milímetros. A célula diz a
 * fração dela da largura; a de uma coluna só dá a largura da coluna, e a mesclada
 * reparte o que sobra entre as colunas que ninguém mediu.
 */
function gridWidths(grid: { columns: number; rows: (GridSlot | null)[][] }, width: number): number[] {
  const known: (number | null)[] = Array.from({ length: grid.columns }, () => null)
  const spans: Array<{ column: number; span: number; width: number }> = []
  for (const slots of grid.rows) {
    let column = 0
    for (const slot of slots) {
      if (slot === null) {
        column++
        continue
      }
      const fraction = (slot.node.attrs?.['cell'] as { width?: number } | undefined)?.width ?? 0
      if (slot.colspan === 1 && known[column] === null && fraction > 0) known[column] = fraction * width
      else if (fraction > 0) spans.push({ column, span: slot.colspan, width: fraction * width })
      column += slot.colspan
    }
  }
  for (const { column, span, width: total } of spans) {
    const range = known.slice(column, column + span)
    const missing = range.filter((value) => value === null).length
    if (missing === 0) continue
    const rest = total - range.reduce<number>((sum, value) => sum + (value ?? 0), 0)
    for (let index = column; index < column + span; index++) {
      if (known[index] === null) known[index] = Math.max(1, rest / missing)
    }
  }
  const measured = known.reduce<number>((sum, value) => sum + (value ?? 0), 0)
  const unknown = known.filter((value) => value === null).length
  const share = unknown === 0 ? 0 : Math.max(1, (width - measured) / unknown)
  return known.map((value) => value ?? share)
}

/** O texto da faixa, com `{n}` e `{total}` como os campos de página e de total. */
function bandText(text: string): string {
  return text
    .split(/(\{n\}|\{total\})/)
    .map((part) =>
      part === '{n}'
        ? '<text:page-number text:select-page="current">1</text:page-number>'
        : part === '{total}'
          ? '<text:page-count>1</text:page-count>'
          : odfText(part),
    )
    .join('')
}

const RELATIVE: Readonly<Record<string, string>> = {
  page: 'page',
  margin: 'page-content',
  column: 'paragraph',
  paragraph: 'paragraph',
  character: 'char',
  line: 'paragraph',
}

const WRAP: Readonly<Record<string, string>> = {
  none: 'run-through',
  square: 'parallel',
  tight: 'parallel',
  through: 'parallel',
  topAndBottom: 'none',
}

function annotationName(id: string): string {
  return `__Annotation__${id}`
}

function listLevelXml(depth: number, info: ListInfo): string {
  const level = info.def.levels[info.level] ?? info.def.levels[0]
  const fmt = level?.fmt ?? (info.kind === 'orderedList' ? 'decimal' : 'bullet')
  const indent = info.indentMm
  const alignment =
    '<style:list-level-properties text:list-level-position-and-space-mode="label-alignment">' +
    `<style:list-level-label-alignment text:label-followed-by="listtab"${attr('text:list-tab-stop-position', mm(indent))}` +
    `${attr('fo:text-indent', mm(-info.hangingMm))}${attr('fo:margin-left', mm(indent))}/>` +
    '</style:list-level-properties>'
  if (fmt === 'bullet') {
    const char = [...(level?.text ?? '')][0] ?? '•'
    return `<text:list-level-style-bullet${attr('text:level', depth)}${attr('text:bullet-char', char)}>${alignment}</text:list-level-style-bullet>`
  }
  const text = level?.text ?? '%1.'
  const placeholders = [...text.matchAll(/%(\d)/g)]
  const firstAt = placeholders[0]?.index ?? text.length
  const last = placeholders[placeholders.length - 1]
  const prefix = placeholders.length === 0 ? text : text.slice(0, firstAt)
  const suffix = last === undefined ? '' : text.slice(last.index + last[0].length)
  const shown = Math.max(1, Math.min(placeholders.length, depth))
  return (
    `<text:list-level-style-number${attr('text:level', depth)}${attr('style:num-prefix', prefix)}${attr('style:num-suffix', suffix)}` +
    ` style:num-format="${placeholders.length === 0 ? '' : xml(numberFormat(fmt))}"` +
    `${shown > 1 ? attr('text:display-levels', shown) : ''}${attr('text:start-value', level?.start ?? 1)}>${alignment}</text:list-level-style-number>`
  )
}

interface GridSlot {
  readonly node: DocumentNode
  readonly colspan: number
  readonly rowspan: number
}

/**
 * A grade da tabela: o ODF quer uma posição por coluna em toda linha, com
 * `covered-table-cell` onde a célula de cima (ou da esquerda) se estende. A
 * célula estendida para a direita leva as cobertas dela junto (`cell`); aqui
 * só entram as cobertas pelas linhas de cima (`null`).
 */
function tableGrid(rows: readonly DocumentNode[]): { columns: number; rows: (GridSlot | null)[][] } {
  const below: number[] = []
  const out: (GridSlot | null)[][] = []
  let columns = 0
  for (const row of rows) {
    const slots: (GridSlot | null)[] = []
    let column = 0
    const take = (): void => {
      while ((below[column] ?? 0) > 0) {
        below[column]!--
        slots.push(null)
        column++
      }
    }
    for (const cell of row.content ?? []) {
      take()
      const colspan = Math.max(1, Math.round(finite(cell.attrs?.['colspan']) ?? 1))
      const rowspan = Math.max(1, Math.round(finite(cell.attrs?.['rowspan']) ?? 1))
      slots.push({ node: cell, colspan, rowspan })
      for (let index = 0; index < colspan; index++) below[column + index] = rowspan - 1
      column += colspan
    }
    take()
    out.push(slots)
    columns = Math.max(columns, column)
  }
  // Toda linha com o mesmo número de posições.
  for (const slots of out) {
    const width = slots.reduce((sum, slot) => sum + (slot === null ? 1 : slot.colspan), 0)
    for (let index = width; index < columns; index++) slots.push(null)
  }
  return { columns, rows: out }
}

/** As larguras em milímetros, da primeira linha, quando cada célula as declara. */
function columnWidthsMm(rows: readonly DocumentNode[], columns: number): number[] | null {
  const first = rows[0]
  if (first === undefined) return null
  const widths: number[] = []
  for (const cell of first.content ?? []) {
    const declared = cell.attrs?.['colwidth']
    if (!Array.isArray(declared) || declared.length === 0) return null
    for (const width of declared) {
      const number = finite(width)
      if (number === null || number <= 0) return null
      widths.push(pxToMm(number))
    }
  }
  return widths.length === columns ? widths : null
}

/** O tamanho da figura: o do documento, o do arquivo, ou um palpite, nessa ordem. */
function sizeMm(
  width: number | null,
  height: number | null,
  picture: Picture,
): { width: number; height: number } {
  const natural = picture.size
  if (width !== null && width > 0 && height !== null && height > 0)
    return { width: pxToMm(width), height: pxToMm(height) }
  if (natural !== null && natural.width > 0 && natural.height > 0) {
    if (width !== null && width > 0)
      return { width: pxToMm(width), height: pxToMm((width * natural.height) / natural.width) }
    return { width: pxToMm(natural.width), height: pxToMm(natural.height) }
  }
  return { width: 40, height: 40 }
}

function plainText(node: DocumentNode): string {
  if (node.type === 'text') return node.text ?? ''
  if (node.type === 'hardBreak') return '\n'
  return (node.content ?? []).map(plainText).join('')
}

function metaXml(model: OdtModel, options: OdtExportOptions): string {
  const properties = model.properties ?? {}
  const text = (element: string, value: string | undefined): string => {
    const trimmed = value?.trim()
    return trimmed === undefined || trimmed === '' ? '' : `<${element}>${xml(trimmed)}</${element}>`
  }
  const date = (element: string, value: string | undefined): string =>
    value !== undefined && /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/.test(value)
      ? `<${element}>${xml(value)}</${element}>`
      : ''
  const keywords = (properties.keywords ?? '')
    .split(/[;,]/)
    .map((word) => word.trim())
    .filter((word) => word !== '')
    .map((word) => `<meta:keyword>${xml(word)}</meta:keyword>`)
    .join('')
  const userDefined = (name: string, value: string | undefined): string =>
    value === undefined || value.trim() === ''
      ? ''
      : `<meta:user-defined${attr('meta:name', name)}>${xml(value.trim())}</meta:user-defined>`
  const revision = finite(properties.revision)
  const minutes = properties.totalTime
  return (
    `${XML_HEAD}<office:document-meta ${NAMESPACES} office:version="1.3"><office:meta>` +
    `<meta:generator>${xml(options.generator ?? 'Librevia')}</meta:generator>` +
    text('dc:title', properties.title) +
    text('dc:subject', properties.subject) +
    text('dc:description', properties.description) +
    keywords +
    text('meta:initial-creator', properties.creator) +
    text('dc:creator', properties.lastModifiedBy || properties.creator) +
    date('meta:creation-date', properties.created) +
    date('dc:date', properties.modified) +
    (revision !== null && revision > 0
      ? `<meta:editing-cycles>${Math.round(revision)}</meta:editing-cycles>`
      : '') +
    (minutes !== undefined && Number.isFinite(minutes) && minutes > 0
      ? `<meta:editing-duration>PT${Math.round(minutes)}M</meta:editing-duration>`
      : '') +
    userDefined('Category', properties.category) +
    userDefined('Company', properties.company) +
    userDefined('Manager', properties.manager) +
    '</office:meta></office:document-meta>'
  )
}

function manifestXml(pictures: readonly Picture[], formulas: readonly Formula[]): string {
  const entry = (path: string, type: string): string =>
    `<manifest:file-entry${attr('manifest:full-path', path)}${attr('manifest:media-type', type)}/>`
  return (
    `${XML_HEAD}<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3">` +
    `<manifest:file-entry manifest:full-path="/" manifest:version="1.3"${attr('manifest:media-type', MIMETYPE)}/>` +
    entry('content.xml', 'text/xml') +
    entry('styles.xml', 'text/xml') +
    entry('meta.xml', 'text/xml') +
    pictures.map((picture) => entry(picture.path, picture.mime)).join('') +
    formulas
      .map(
        (formula) =>
          `<manifest:file-entry${attr('manifest:full-path', `${formula.path}/`)} manifest:version="1.3"` +
          `${attr('manifest:media-type', FORMULA_MIMETYPE)}/>` +
          entry(`${formula.path}/content.xml`, 'text/xml'),
      )
      .join('') +
    '</manifest:manifest>'
  )
}
