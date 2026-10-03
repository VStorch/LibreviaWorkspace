/**
 * Modelo canônico do documento.
 *
 * `doc` é o JSON do ProseMirror — o mesmo que o Tiptap edita nativamente — e
 * `page` carrega o que não cabe no fluxo de texto. A separação existe porque o
 * DOCX é lido para estas duas partes, e o PDF é gerado a partir delas.
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
  /**
   * Cabeçalho e rodapé em texto simples, digitados pelo usuário. Aceitam `{n}`
   * para o número da página e `{total}` para o total — a substituição acontece
   * na hora de gerar o PDF, que é quando o número de páginas passa a existir.
   */
  readonly header: string
  readonly footer: string
  /**
   * O cabeçalho real do documento importado, quando existe. **Manda na
   * exibição**: um `.docx` corporativo traz logotipo e numeração que o campo
   * de texto acima não representaria.
   */
  readonly headerBand: Band | null
  readonly footerBand: Band | null
  /**
   * Faixas de primeira página e de páginas pares.
   *
   * Só existem quando o documento **liga** os interruptores correspondentes
   * (`w:titlePg`, `w:evenAndOddHeaders`). O Word guarda as partes mesmo com eles
   * desligados, e usá-las sem conferir poria o cabeçalho da capa em todas as
   * páginas — ver `PageReader.HasTitlePage`.
   */
  readonly firstHeaderBand: Band | null
  readonly firstFooterBand: Band | null
  readonly evenHeaderBand: Band | null
  readonly evenFooterBand: Band | null
  /**
   * Distância da faixa à borda do papel, em milímetros (`w:pgMar/@header`).
   *
   * É a origem vertical das âncoras de dentro do cabeçalho: elas se dizem
   * relativas ao parágrafo, e o parágrafo do cabeçalho começa aqui.
   */
  readonly headerDistanceMm: number
  readonly footerDistanceMm: number
  /**
   * Numeração de página (`w:pgNumType`): o formato e o número da primeira folha.
   *
   * Opcionais, como os interruptores abaixo: o `.sdoc` gravado antes não os
   * tem, e ausência quer dizer "decimal, a partir de 1" na tela e "não mexa" na
   * gravação.
   */
  readonly pageNumberFormat?: PageNumberFormat | undefined
  readonly pageNumberStart?: number | null | undefined
  /**
   * "Primeira página diferente" (`w:titlePg`) e "Pares e ímpares diferentes"
   * (`w:evenAndOddHeaders`). Ausentes, valem pelo que as faixas dizem: um
   * rascunho de antes só trazia a faixa da capa quando o interruptor estava
   * ligado.
   */
  readonly titlePage?: boolean | null | undefined
  readonly evenAndOddHeaders?: boolean | null | undefined
  /**
   * Como a seção começa (`w:sectPr/w:type`). Ausente (rascunho de antes das
   * seções) vale "próxima página" na tela e "não mexa" na gravação.
   */
  readonly start?: SectionStart | undefined
  /**
   * As colunas da seção (`w:cols`). Ausente (rascunho de antes) é uma coluna na
   * tela e "não mexa" na gravação.
   */
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
 * Uma seção antes da última.
 *
 * O `id` é o elo com o texto: o parágrafo que **encerra** a seção leva o mesmo
 * valor no atributo `sectionBreak`, como no OOXML o `w:sectPr` mora no parágrafo
 * que fecha a seção. A última seção é `DocumentModel.page` — a do corpo, que não
 * tem parágrafo.
 *
 * Cada seção traz só as faixas que **declara**; nula, da segunda em diante, é
 * "vincular ao anterior" — ver `effectiveSections`.
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
 * Nó do ProseMirror em forma serializável.
 *
 * As coleções são mutáveis de propósito: este tipo precisa ser aceito onde o
 * Tiptap espera `JSONContent`, e um array `readonly` não é atribuível a um
 * array comum. As propriedades continuam `readonly` — o que importa é não
 * reescrever o nó por engano.
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
  /**
   * As seções antes da última, em ordem. Ausente é documento de uma seção. Fora
   * dos nós pelo mesmo motivo dos estilos: mudar o papel de uma seção não pode
   * fazer o parágrafo da marca parecer editado.
   */
  readonly sections?: readonly SectionSetup[]
  readonly doc: DocumentNode
  /**
   * Os estilos do documento — **fora dos nós**, e é isso que os torna seguros.
   *
   * A gravação cirúrgica decide o que preservar comparando a impressão digital de
   * cada bloco com a que o leitor produziu, e ela é feita do que está dentro do
   * nó. Um estilo guardado ali faria todo bloco parecer mudado, e o documento
   * inteiro seria reescrito — exatamente o que este projeto existe para evitar.
   *
   * Ver `styles.ts`: é deles que nasce o CSS da tela e do PDF (`style-css.ts`).
   */
  readonly styles: StyleSheet
  /**
   * Os blocos vieram **achatados**: cada um com a formatação efetiva — padrões,
   * estilo e direta —, como o leitor os produzia antes de levar só a direta.
   *
   * É o rascunho gravado por uma versão anterior (formato `.sdoc` < 4). Na tela
   * não faz diferença — o inline vence a regra do estilo, e o achatado já diz
   * tudo —, mas na gravação faz: os blocos são comparados com uma leitura do
   * original, e ela tem de ser achatada também, senão todo bloco pareceria mudado
   * e o documento inteiro seria reescrito. Ausente é falso.
   */
  readonly flattened?: boolean
  /**
   * O rascunho é de antes das **referências** (formato `.sdoc` < 5): os nós não
   * trazem marcador, campo, link interno nem sumário, que o leitor passou a
   * produzir na versão 5. Mesmo motivo de `flattened`: a gravação compara com
   * uma leitura do original feita como era então. Ausente é falso.
   */
  readonly beforeReferences?: boolean
  /**
   * O rascunho é de antes das **seções** (formato `.sdoc` < 6): os parágrafos
   * que encerram seção não trazem `sectionBreak`. Mesmo motivo de `flattened`.
   */
  readonly beforeSections?: boolean
  /**
   * Os marcadores do arquivo que não viraram nó — entre linhas de tabela, soltos
   * entre blocos, no cabeçalho ou numa caixa de texto. Existem, e continuam no
   * arquivo; a referência que os cita não está quebrada, e "Atualizar campos"
   * deixa o resultado dela como o Word o deixou.
   */
  readonly outsideBookmarks?: readonly string[]
  /**
   * Os comentários do arquivo. O corpo mora aqui, fora dos nós, pelo mesmo
   * motivo dos estilos; no texto ficam só as pontas da âncora (`commentStart` e
   * `commentEnd`, uma por conversa — a resposta não tem nó). Só os que o texto
   * sustenta (`resolveComments`); o que não mudou volta ao arquivo byte a byte.
   */
  readonly comments?: readonly DocumentComment[]
  /**
   * O rascunho é de antes dos **comentários** (formato `.sdoc` < 7): os nós não
   * trazem a âncora. Mesmo motivo de `flattened`.
   */
  readonly beforeComments?: boolean
  /**
   * O documento grava controlando alterações — o `w:trackRevisions` do arquivo
   *. Ausente é "não mexa": o arquivo fica como está.
   */
  readonly trackChanges?: boolean
  /**
   * O rascunho é de antes das **revisões** (formato `.sdoc` < 8): os nós não
   * trazem as marcas `insertion`/`deletion` nem a revisão de bloco. Mesmo motivo
   * de `flattened`.
   */
  readonly beforeRevisions?: boolean
  /**
   * Como o documento numera as notas de rodapé e as de fim. Fora dos nós pelo
   * mesmo motivo dos estilos; a referência (`noteRef`) não guarda número — ele
   * é a ordem dela no documento. Ausente é a numeração do Word: 1, 2, 3 nas de
   * rodapé e i, ii, iii nas de fim.
   */
  readonly notes?: DocumentNotes
  /**
   * O rascunho é de antes das **notas** (formato `.sdoc` < 9): os nós não trazem
   * o `noteRef`. Mesmo motivo de `flattened`.
   */
  readonly beforeNotes?: boolean
  /**
   * O rascunho é de antes das **equações** (formato `.sdoc` < 11): os nós não
   * trazem o `math`, e a equação ficava escondida no parágrafo. Mesmo motivo de
   * `flattened`.
   */
  readonly beforeMath?: boolean
  /**
   * As propriedades do documento — `docProps/core.xml` e parte de
   * `docProps/app.xml`. Fora dos nós pelo mesmo motivo dos estilos.
   *
   * Na gravação em DOCX cada campo é um **remendo**: ausente é "deixe como está
   * no arquivo", e a cadeia vazia é "apague". Por isso o rascunho de antes delas
   * (`.sdoc` < 10) não precisa de marca: sem `properties`, nada em `docProps/`
   * é tocado, e o pacote de origem guarda as do arquivo byte a byte.
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

/** Largura útil do texto: é ela que define a medida da moldura na tela. */
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
 * Onde a coluna de texto começa e termina na folha.
 *
 * A margem é um piso, não uma posição. Quando o cabeçalho é mais alto do que a
 * distância dele até a borda mais a margem de cima — e o cabeçalho corporativo
 * em grade quase sempre é — o Word e o LibreOffice **descem o corpo** até
 * debaixo dele. Sem isso a primeira linha do texto era escrita por cima da
 * última do cabeçalho, e o mesmo encontro acontecia no pé com o rodapé.
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

/** O caminho de volta, para o que foi medido na tela e vai ser posto em mm. */
export function pxToMm(px: number): number {
  return (px * 25.4) / 96
}

/**
 * Margem válida é a que deixa espaço útil.
 *
 * Sem este limite o usuário consegue pedir margens que somam mais que a página
 * — e o resultado seria uma área de texto de largura negativa.
 */
export function isValidMargins(page: PageSetup): boolean {
  const { width, height } = pageDimensionsMm(page)
  const values = [page.margins.top, page.margins.right, page.margins.bottom, page.margins.left]

  if (values.some((value) => !Number.isFinite(value) || value < 0)) return false
  return page.margins.left + page.margins.right < width && page.margins.top + page.margins.bottom < height
}
