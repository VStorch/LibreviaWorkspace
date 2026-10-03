/**
 * `doc` é o JSON do ProseMirror, o mesmo que o Tiptap edita; `page` e o resto
 * carregam o que não cabe no fluxo de texto.
 */

import { NO_BANDS, type Band, type BandHeights } from './band.js'
import { BUILTIN_STYLES, type StyleSheet } from './styles.js'

export const PageSize = {
  A4: 'A4',
  Letter: 'Letter',
} as const
export type PageSize = (typeof PageSize)[keyof typeof PageSize]

export const PageOrientation = {
  Portrait: 'portrait',
  Landscape: 'landscape',
} as const
export type PageOrientation = (typeof PageOrientation)[keyof typeof PageOrientation]

/** Margens em milímetros — a unidade que aparece na interface. */
export interface Margins {
  readonly top: number
  readonly right: number
  readonly bottom: number
  readonly left: number
}

export interface PageSetup {
  readonly size: PageSize
  readonly orientation: PageOrientation
  readonly margins: Margins
  /** Texto simples digitado pela pessoa; `{n}` e `{total}` são trocados ao gerar o PDF. */
  readonly header: string
  readonly footer: string
  /** O cabeçalho do documento importado; quando existe, manda na exibição. */
  readonly headerBand: Band | null
  readonly footerBand: Band | null
  /**
   * Só existem quando o documento liga `w:titlePg` ou `w:evenAndOddHeaders`: o
   * Word guarda as partes mesmo com eles desligados — ver `PageReader.HasTitlePage`.
   */
  readonly firstHeaderBand: Band | null
  readonly firstFooterBand: Band | null
  readonly evenHeaderBand: Band | null
  readonly evenFooterBand: Band | null
  /** `w:pgMar/@header`: a origem vertical das âncoras de dentro do cabeçalho. */
  readonly headerDistanceMm: number
  readonly footerDistanceMm: number
  /** `w:pgNumType`. Ausente vale "decimal, a partir de 1" na tela e "não mexa" na gravação. */
  readonly pageNumberFormat?: PageNumberFormat | undefined
  readonly pageNumberStart?: number | null | undefined
  /** `w:titlePg` e `w:evenAndOddHeaders`. Ausentes, valem pelo que as faixas dizem. */
  readonly titlePage?: boolean | null | undefined
  readonly evenAndOddHeaders?: boolean | null | undefined
  /** `w:sectPr/w:type`. Ausente vale "próxima página" na tela e "não mexa" na gravação. */
  readonly start?: SectionStart | undefined
  /** `w:cols`. Ausente vale uma coluna na tela e "não mexa" na gravação. */
  readonly columns?: SectionColumns | undefined
}

/** `w:cols`: quantas colunas, o espaço entre elas (mm) e a linha separadora. */
export interface SectionColumns {
  readonly count: number
  readonly spaceMm: number
  readonly separator: boolean
  /**
   * Larguras diferentes, como o arquivo as declara (`w:equalWidth="0"`). A tela
   * desenha colunas iguais; mudar as colunas no painel as iguala.
   */
  readonly widthsMm?: number[] | undefined
}

/** Os começos de seção do OOXML — os de `w:type/@w:val`. */
export const SECTION_STARTS = ['nextPage', 'continuous', 'evenPage', 'oddPage', 'nextColumn'] as const
export type SectionStart = (typeof SECTION_STARTS)[number]

/**
 * O parágrafo que **encerra** a seção leva o mesmo `id` em `sectionBreak`, como
 * no OOXML o `w:sectPr` mora no parágrafo que fecha a seção. A última seção é
 * `DocumentModel.page`. Faixa nula, da segunda seção em diante, é "vincular ao
 * anterior" — ver `effectiveSections`.
 */
export interface SectionSetup extends PageSetup {
  readonly id: string
}

/** Os formatos de número de página que o painel oferece — os de `w:pgNumType/@w:fmt`. */
export const PAGE_NUMBER_FORMATS = [
  'decimal',
  'lowerRoman',
  'upperRoman',
  'lowerLetter',
  'upperLetter',
] as const
export type PageNumberFormat = (typeof PAGE_NUMBER_FORMATS)[number]

/**
 * Coleções mutáveis porque o Tiptap espera `JSONContent`, e um array `readonly`
 * não é atribuível a um comum.
 */
export interface DocumentNode {
  readonly type: string
  readonly content?: DocumentNode[]
  readonly text?: string
  readonly attrs?: Record<string, unknown>
  readonly marks?: { readonly type: string; readonly attrs?: Record<string, unknown> }[]
}

export interface DocumentModel {
  /** A última seção — a do corpo, e a única do documento de uma seção só. */
  readonly page: PageSetup
  /** Antes da última, em ordem. Fora dos nós pelo mesmo motivo de `styles`. */
  readonly sections?: readonly SectionSetup[]
  readonly doc: DocumentNode
  /**
   * Fora dos nós: a impressão digital de um bloco é feita do que está dentro
   * dele, e um estilo ali faria todo bloco parecer mudado na gravação cirúrgica.
   */
  readonly styles: StyleSheet
  /**
   * Rascunho `.sdoc` < 4, com a formatação efetiva em cada bloco. A gravação o
   * compara com uma leitura achatada do original, senão todo bloco pareceria mudado.
   */
  readonly flattened?: boolean
  /** Rascunho `.sdoc` < 5: sem marcador, campo, link interno nem sumário nos nós. Mesmo motivo de `flattened`. */
  readonly beforeReferences?: boolean
  /** Rascunho `.sdoc` < 6: sem `sectionBreak`. Mesmo motivo de `flattened`. */
  readonly beforeSections?: boolean
  /**
   * Marcadores do arquivo que não viraram nó — entre linhas de tabela, soltos
   * entre blocos, no cabeçalho ou numa caixa. A referência que os cita não está
   * quebrada, e "Atualizar campos" deixa o resultado dela como o Word deixou.
   */
  readonly outsideBookmarks?: readonly string[]
  /**
   * O corpo mora fora dos nós, como `styles`; no texto ficam só as pontas da
   * âncora, uma por conversa. Só os que o texto sustenta (`resolveComments`).
   */
  readonly comments?: readonly DocumentComment[]
  /** Rascunho `.sdoc` < 7: sem a âncora nos nós. Mesmo motivo de `flattened`. */
  readonly beforeComments?: boolean
  /** `w:trackRevisions`. Ausente é "não mexa". */
  readonly trackChanges?: boolean
  /** Rascunho `.sdoc` < 8: sem as marcas de revisão. Mesmo motivo de `flattened`. */
  readonly beforeRevisions?: boolean
  /**
   * Fora dos nós, como `styles`: o número de uma nota é a ordem da referência.
   * Ausente é a numeração do Word: 1, 2, 3 nas de rodapé e i, ii, iii nas de fim.
   */
  readonly notes?: DocumentNotes
  /** Rascunho `.sdoc` < 9: sem `noteRef`. Mesmo motivo de `flattened`. */
  readonly beforeNotes?: boolean
  /** Rascunho `.sdoc` < 11: sem `math`. Mesmo motivo de `flattened`. */
  readonly beforeMath?: boolean
  /**
   * `docProps/core.xml` e parte de `app.xml`, fora dos nós. Na gravação cada
   * campo é um remendo: ausente é "deixe como está", vazio é "apague".
   */
  readonly properties?: DocumentProperties
}

/**
 * As propriedades de um documento, como o Word as mostra em Arquivo →
 * Propriedades. As datas são W3CDTF (`2026-10-02T12:00:00Z`), como no pacote.
 */
export interface DocumentProperties {
  readonly title?: string
  readonly subject?: string
  /** `dc:creator`: o(s) autor(es), separados por ponto e vírgula, como no Word. */
  readonly creator?: string
  readonly keywords?: string
  readonly category?: string
  /** `dc:description`: o que o Word chama de Comentários. */
  readonly description?: string
  readonly lastModifiedBy?: string
  readonly revision?: string
  readonly created?: string
  readonly modified?: string
  /** `docProps/app.xml`. */
  readonly company?: string
  readonly manager?: string
  /** `TotalTime` do `app.xml`, em minutos. Só lido: o editor não o mede. */
  readonly totalTime?: number
}

/** A numeração de um tipo de nota, como `w:footnotePr`/`w:endnotePr` a descrevem. */
export interface NoteNumbering {
  /** `decimal`, `lowerRoman`, `upperLetter`, `chicago`… */
  readonly numFmt?: string
  readonly start?: number
  /** `continuous`, `eachSect`, `eachPage`. */
  readonly restart?: string
  readonly pos?: string
}

export interface DocumentNotes {
  readonly footnotePr?: NoteNumbering
  readonly endnotePr?: NoteNumbering
}

/** Um comentário, como `word/comments.xml` e `word/commentsExtended.xml` o descrevem. */
export interface DocumentComment {
  /** O `w:id` — o mesmo `cid` das pontas no texto. */
  readonly id: string
  /** O comentário que este responde. Ausente é o que abre a conversa. */
  readonly parentId?: string
  readonly author: string
  readonly initials?: string
  /** Como o arquivo o traz (ISO 8601); vazio quando não traz. */
  readonly date: string
  /** O texto de cada parágrafo, sem formatação. */
  readonly paragraphs: readonly string[]
  /** Resolvido (`w15:done`) — vale para a conversa, pelo comentário que a abre. */
  readonly done: boolean
  readonly paraId?: string
  /** O corpo tem formatação, imagem ou campo que o texto simples não mostra. */
  readonly rich?: boolean
}

export const PAGE_DIMENSIONS_MM: Record<PageSize, { width: number; height: number }> = {
  [PageSize.A4]: { width: 210, height: 297 },
  [PageSize.Letter]: { width: 216, height: 279 },
}

/** Equivalente ao padrão "Normal" do Word: 2,54 cm em volta. */
export const DEFAULT_PAGE_SETUP: PageSetup = {
  size: PageSize.A4,
  orientation: PageOrientation.Portrait,
  margins: { top: 25, right: 25, bottom: 25, left: 25 },
  header: '',
  footer: '',
  headerBand: null,
  footerBand: null,
  firstHeaderBand: null,
  firstFooterBand: null,
  evenHeaderBand: null,
  evenFooterBand: null,
  headerDistanceMm: 12.5,
  footerDistanceMm: 12.5,
}

export const EMPTY_DOCUMENT: DocumentNode = {
  type: 'doc',
  content: [{ type: 'paragraph' }],
}

export function createEmptyDocument(): DocumentModel {
  return { page: DEFAULT_PAGE_SETUP, doc: EMPTY_DOCUMENT, styles: BUILTIN_STYLES }
}

/** Largura e altura já considerando a orientação. */
export function pageDimensionsMm(page: PageSetup): { width: number; height: number } {
  const base = PAGE_DIMENSIONS_MM[page.size]
  return page.orientation === PageOrientation.Landscape
    ? { width: base.height, height: base.width }
    : { width: base.width, height: base.height }
}

/** A largura que define a moldura na tela. */
export function contentWidthMm(page: PageSetup): number {
  const { width } = pageDimensionsMm(page)
  return width - page.margins.left - page.margins.right
}

export function contentHeightMm(page: PageSetup, bands: BandHeights = NO_BANDS): number {
  const { height } = pageDimensionsMm(page)
  const inset = contentInsetsMm(page, bands)
  return height - inset.top - inset.bottom
}

/**
 * A margem é um piso: quando o cabeçalho é mais alto que a distância dele até a
 * borda mais a margem, o Word e o LibreOffice descem o corpo até debaixo dele.
 */
export function contentInsetsMm(page: PageSetup, bands: BandHeights): { top: number; bottom: number } {
  return {
    top: Math.max(page.margins.top, page.headerDistanceMm + bands.headerMm),
    bottom: Math.max(page.margins.bottom, page.footerDistanceMm + bands.footerMm),
  }
}

/** Conversão CSS: 1 polegada = 96 px = 25,4 mm. */
export function mmToPx(mm: number): number {
  return (mm * 96) / 25.4
}

export function pxToMm(px: number): number {
  return (px * 25.4) / 96
}

/** Margens que somam mais que a página dariam área de texto negativa. */
export function isValidMargins(page: PageSetup): boolean {
  const { width, height } = pageDimensionsMm(page)
  const values = [page.margins.top, page.margins.right, page.margins.bottom, page.margins.left]

  if (values.some((value) => !Number.isFinite(value) || value < 0)) return false
  return page.margins.left + page.margins.right < width && page.margins.top + page.margins.bottom < height
}
