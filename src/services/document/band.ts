import type { DocumentNode, PageSetup } from './model.js'
import { formatNumber } from './list-numbering.js'
import type { FloatingObject } from './floating.js'

/**
 * O cabeçalho ou rodapé preservado do documento. O texto das peças com endereço
 * é editável e volta para o `w:t` de onde veio; o resto da parte OOXML volta
 * intacto.
 */
export interface Band {
  readonly left: BandPiece[]
  readonly center: BandPiece[]
  readonly right: BandPiece[]
  readonly rule: boolean
  /** O que não cabe em três colunas: desenho com posição de verdade, que pode vir girado. */
  readonly floats: FloatingObject[]
  /** A grade, quando o cabeçalho é uma tabela: logotipo em célula mesclada, título ao lado. */
  readonly rows: BandRow[]
}

export interface BandRow {
  readonly cells: BandCell[]
}

/** Uma célula da grade: o que está escrito nela e o retângulo que ela ocupa. */
export interface BandCell {
  readonly pieces: BandPiece[]
  /** Fração da largura da grade, de 0 a 1. */
  readonly width: number
  readonly span: number
  readonly rowSpan: number
  readonly align?: string | undefined
  /** Iniciais dos lados com risco: `t`, `l`, `b`, `r`. */
  readonly borders: string
}

export interface BandPiece {
  readonly kind: 'text' | 'image' | 'pageNumber' | 'totalPages'
  // `| undefined` explícito por causa de `exactOptionalPropertyTypes`: este
  // tipo precisa ser atribuível ao que o zod infere no schema compartilhado.
  readonly text?: string | undefined
  readonly src?: string | undefined
  readonly width?: number | undefined
  readonly height?: number | undefined
  readonly bold: boolean
  readonly italic: boolean
  readonly color?: string | undefined
  readonly fontSize?: string | undefined
  /** Pilha de CSS, como o leitor a resolveu. */
  readonly fontFamily?: string | undefined
  /** Cabeçalho e rodapé são parágrafos, e cada parágrafo é uma linha. */
  readonly line?: boolean | undefined
  /**
   * Onde a peça mora no arquivo: a gravação escreve no `w:t` dela e não toca o
   * resto do cabeçalho. Número de página, imagem e tabulação não têm endereço.
   */
  readonly pid?: string | undefined
  /**
   * O texto do arquivo já trazia `{n}` ou `{total}` escritos. É texto: a tela não
   * o troca pelo número, e a gravação não o transforma em campo.
   */
  readonly literal?: boolean | undefined
}

/** Compartilhada pela tela e pelo papel, para os dois desenhos não divergirem. */
export function linesOf(pieces: readonly BandPiece[]): BandPiece[][] {
  const lines: BandPiece[][] = []
  for (const piece of pieces) {
    if (piece.line === true || lines.length === 0) lines.push([])
    lines[lines.length - 1]!.push(piece)
  }

  return lines
}

/** Medida na folha: só existe depois de desenhar. */
export interface BandHeights {
  readonly headerMm: number
  readonly footerMm: number
}

export const NO_BANDS: BandHeights = { headerMm: 0, footerMm: 0 }

/**
 * A ordem é a do Word: a capa manda sobre a paridade, e a paridade sobre o
 * padrão. Faixa ausente cai no padrão; `hasBandContent` decide se a folha fica
 * limpa.
 */
export function bandForPage(page: PageSetup, sheet: number, kind: 'header' | 'footer'): Band | null {
  const first = kind === 'header' ? page.firstHeaderBand : page.firstFooterBand
  const even = kind === 'header' ? page.evenHeaderBand : page.evenFooterBand
  const fallback =
    (kind === 'header' ? page.headerBand : page.footerBand) ??
    plainBand(kind === 'header' ? page.header : page.footer)

  // Ligado sem faixa própria quer dizer folha limpa, como no Word: é o uso de
  // "Primeira página diferente" na capa sem número.
  if (sheet === 1 && usesTitlePage(page)) return first
  // A paridade é a do número impresso, e não a da folha: começando em 2, a
  // primeira folha já é par.
  if (usesEvenAndOdd(page) && pageNumberOf(page, sheet) % 2 === 0) return even
  return fallback
}

/** "Primeira página diferente": o que o documento diz, ou o que as faixas dão a entender. */
export function usesTitlePage(page: PageSetup): boolean {
  return page.titlePage ?? (page.firstHeaderBand !== null || page.firstFooterBand !== null)
}

/** "Pares e ímpares diferentes", pelo mesmo critério. */
export function usesEvenAndOdd(page: PageSetup): boolean {
  return page.evenAndOddHeaders ?? (page.evenHeaderBand !== null || page.evenFooterBand !== null)
}

/** O número impresso na folha `sheet` (a contar de 1): o início de `w:pgNumType` mais o avanço. */
export function pageNumberOf(page: PageSetup, sheet: number): number {
  return (page.pageNumberStart ?? 1) + sheet - 1
}

/** O número da folha como o campo `PAGE` o escreve, no formato de `w:pgNumType`. */
export function pageLabel(page: PageSetup, sheet: number): string {
  return formatNumber(pageNumberOf(page, sheet), page.pageNumberFormat ?? 'decimal')
}

/** O texto pode trazer `{n}` e `{total}` até a gravação os transformar em campo (`BandWriter.Rewrite`). */
export function pieceText(piece: BandPiece, label: string, total: number): string {
  if (piece.kind === 'pageNumber') return label
  if (piece.kind === 'totalPages') return String(total)
  if (piece.literal === true) return piece.text ?? ''
  return substituteFields(piece.text ?? '', label, total)
}

export function substituteFields(text: string, label: string, total: number): string {
  return text.replaceAll('{n}', label).replaceAll('{total}', String(total))
}

/** Fonte da linha de texto simples: a de `TemplateStyles.BandFont`, que é a que o arquivo recebe. */
const PLAIN_BAND_FONT = 'Calibri, Carlito, sans-serif'

/**
 * A linha de "Configurar página" como faixa, para a folha paginada a desenhar
 * como o arquivo a grava: centralizada, em 9 pt cinza.
 */
export function plainBand(text: string): Band | null {
  if (text.trim().length === 0) return null
  const style = { bold: false, italic: false, color: '#444444', fontSize: '9pt', fontFamily: PLAIN_BAND_FONT }
  const center: BandPiece[] = []
  for (const part of text.split(/(\{n\}|\{total\})/)) {
    if (part === '') continue
    if (part === '{n}') center.push({ kind: 'pageNumber', ...style })
    else if (part === '{total}') center.push({ kind: 'totalPages', ...style })
    else center.push({ kind: 'text', text: part, ...style })
  }
  return { left: [], center, right: [], rule: false, floats: [], rows: [] }
}

/**
 * No arquivo o cabeçalho é um só: trocar a peça atualiza todas as folhas, como
 * no Word. Devolve a mesma configuração quando nada muda, para não sujar o
 * documento.
 */
export function editBandPiece<T extends PageSetup>(page: T, pid: string, text: string): T {
  let changed = false

  const inPieces = (pieces: BandPiece[]): BandPiece[] =>
    pieces.map((piece) => {
      if (piece.pid !== pid || piece.text === text) return piece
      changed = true
      return { ...piece, text }
    })

  const updated = mapBands(page, (band) => ({
    ...band,
    left: inPieces(band.left),
    center: inPieces(band.center),
    right: inPieces(band.right),
    rows: band.rows.map((row) => ({
      cells: row.cells.map((cell) => ({ ...cell, pieces: inPieces(cell.pieces) })),
    })),
  }))

  return changed ? updated : page
}

/** A caixa vem inteira: digitar dentro dela abre e fecha parágrafos. */
export function editBandFloat<T extends PageSetup>(page: T, bid: string, content: DocumentNode[]): T {
  let changed = false

  const updated = mapBands(page, (band) => ({
    ...band,
    floats: band.floats.map((object) => {
      if (object.bid !== bid) return object
      if (JSON.stringify(object.content ?? []) === JSON.stringify(content)) return object
      changed = true
      return { ...object, content }
    }),
  }))

  return changed ? updated : page
}

export function hasBandContent(band: Band | null): band is Band {
  return (
    band !== null &&
    (band.left.length > 0 ||
      band.center.length > 0 ||
      band.right.length > 0 ||
      band.rows.length > 0 ||
      band.rule)
  )
}

/** As seis faixas: capa, pares e padrão, para cabeçalho e rodapé. */
function mapBands<T extends PageSetup>(page: T, transform: (band: Band) => Band): T {
  const at = (band: Band | null): Band | null => (band === null ? null : transform(band))

  return {
    ...page,
    headerBand: at(page.headerBand),
    footerBand: at(page.footerBand),
    firstHeaderBand: at(page.firstHeaderBand),
    firstFooterBand: at(page.firstFooterBand),
    evenHeaderBand: at(page.evenHeaderBand),
    evenFooterBand: at(page.evenFooterBand),
  }
}

/** Metade da margem: a faixa é mais larga que a coluna de texto, como no documento corporativo. */
export function bandInsetMm(page: PageSetup): number {
  return Math.min(page.margins.left, page.margins.right) / 2
}

/**
 * Sem o cursor numa faixa, o campo vai para o fim do rodapé, como no Word. Num
 * rodapé do arquivo sem texto editável devolve `null`: inventar um parágrafo na
 * parte do Word é o que a gravação cirúrgica não faz.
 */
export function appendPageField(page: PageSetup, token: '{n}' | '{total}'): PageSetup | null {
  const band = page.footerBand
  if (band === null || !hasBandContent(band)) {
    const footer = page.footer.trimEnd()
    return { ...page, footer: footer === '' ? token : `${footer} ${token}` }
  }

  const candidates = [
    ...band.rows.flatMap((row) => row.cells.flatMap((cell) => cell.pieces)),
    ...band.left,
    ...band.center,
    ...band.right,
  ].filter((piece) => piece.kind === 'text' && piece.pid !== undefined)
  const target = candidates.at(-1)
  if (target === undefined || target.pid === undefined) return null

  const text = target.text ?? ''
  const joined = text === '' || text.endsWith(' ') ? `${text}${token}` : `${text} ${token}`
  return editBandPiece(page, target.pid, joined)
}
