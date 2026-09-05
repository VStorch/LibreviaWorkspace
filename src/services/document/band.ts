/**
 * A faixa: o cabeçalho e o rodapé que vieram do documento.
 *
 * Não é o campo de texto que o usuário digita na configuração de página — esse
 * mora em `PageSetup`. É a parte OOXML **preservada**: três colunas, um filete,
 * uma grade quando o cabeçalho é uma tabela, e os objetos ancorados que não
 * cabem em coluna nenhuma. O texto das peças que trazem endereço é editável e
 * volta para o `w:t` de onde veio; todo o resto a gravação devolve intacto.
 *
 * Vive num módulo próprio porque é isto que ele é: um conceito com regras suas,
 * usado por quem desenha a tela, por quem monta o papel e por quem grava.
 */

import type { DocumentNode, PageSetup } from './model.js'
import type { FloatingObject } from './floating.js'

/**
 * Cabeçalho ou rodapé preservado de um documento do Word.
 *
 * Três colunas e um filete opcional. O texto das peças que trazem endereço é
 * editável, e volta para o `w:t` de onde veio; todo o resto da parte OOXML
 * volta intacto. Ver docs/02-docx-cirurgico.md.
 */
export interface Band {
  readonly left: BandPiece[]
  readonly center: BandPiece[]
  readonly right: BandPiece[]
  readonly rule: boolean
  /**
   * Objetos ancorados da faixa.
   *
   * O que não cabe em três colunas: desenho com posição de verdade, que pode vir
   * girado. A marca lateral do corpus é uma faixa de 28,6 mm **em pé** — não
   * entra numa banda de 10 mm de altura, e achatá-la ali a desenhava deitada.
   */
  readonly floats: FloatingObject[]
  /**
   * A grade, quando o cabeçalho é uma tabela.
   *
   * A outra metade do cabeçalho corporativo, e a que não cabe em três colunas:
   * logotipo numa célula mesclada por quatro linhas, título ao lado, numeração
   * à direita. Espalhada por esquerda, centro e direita ela virava uma sopa de
   * palavras que ainda por cima transbordava sobre o texto.
   */
  readonly rows: BandRow[]
}

/** Uma linha da grade do cabeçalho. */
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

/** Um pedaço de cabeçalho vindo do documento: texto, imagem ou campo. */
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
  /**
   * A peça abre linha nova.
   *
   * Cabeçalho e rodapé são feitos de parágrafos, e um parágrafo é uma linha.
   * Sem isto o rodapé de três linhas do modelo de manual saía como uma frase
   * só, emendada na largura da folha.
   */
  readonly line?: boolean | undefined
  /**
   * Onde esta peça mora no arquivo: a relação, o parágrafo e a peça nele.
   *
   * É o que torna a faixa editável sem deixar de ser cirúrgica — a gravação
   * escreve no `w:t` desta peça e não olha para o resto do cabeçalho, que ela
   * não saberia gerar de novo. Peça sem endereço não é editável: número de
   * página, imagem e tabulação não têm texto próprio no arquivo onde escrever.
   */
  readonly pid?: string | undefined
}

/**
 * As peças agrupadas em linhas, quebrando onde o arquivo abre parágrafo.
 *
 * Compartilhada entre a tela e o papel de propósito: dois agrupamentos
 * parecidos escritos em dois arquivos é como os dois desenhos divergem.
 */
export function linesOf(pieces: readonly BandPiece[]): BandPiece[][] {
  const lines: BandPiece[][] = []
  for (const piece of pieces) {
    if (piece.line === true || lines.length === 0) lines.push([])
    lines[lines.length - 1]!.push(piece)
  }

  return lines
}

/**
 * Altura desenhada de cada faixa, medida na folha.
 *
 * É a única parte desta conta que nenhum arquivo diz: um cabeçalho de três
 * linhas ocupa o que a fonte e a quebra derem, e isso só existe depois de
 * desenhar.
 */
export interface BandHeights {
  readonly headerMm: number
  readonly footerMm: number
}

export const NO_BANDS: BandHeights = { headerMm: 0, footerMm: 0 }

/**
 * Qual faixa desenhar nesta página.
 *
 * A ordem é a do Word: a capa manda sobre a paridade, e a paridade sobre o
 * padrão. Faixa ausente cai no padrão, e não em nada — um documento que declara
 * primeira página distinta mas deixa a faixa vazia quer a folha limpa, e é o
 * `hasBandContent` de quem desenha que decide isso.
 */
export function bandForPage(page: PageSetup, pageNumber: number, kind: 'header' | 'footer'): Band | null {
  const first = kind === 'header' ? page.firstHeaderBand : page.firstFooterBand
  const even = kind === 'header' ? page.evenHeaderBand : page.evenFooterBand
  const fallback = kind === 'header' ? page.headerBand : page.footerBand

  if (pageNumber === 1 && first !== null) return first
  if (pageNumber % 2 === 0 && even !== null) return even
  return fallback
}

/**
 * O texto de uma peça da faixa, trocado onde quer que ela esteja.
 *
 * O mesmo cabeçalho é desenhado em todas as folhas, mas no arquivo ele é um só:
 * o endereço é único, e trocá-lo aqui atualiza todas as folhas de uma vez — que
 * é como o Word se comporta quando se edita um cabeçalho.
 *
 * Devolve a mesma configuração quando não há o que trocar, para não sujar o
 * documento por um clique que não mudou nada.
 */
export function editBandPiece(page: PageSetup, pid: string, text: string): PageSetup {
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

/**
 * O conteúdo de uma caixa da faixa, trocado onde quer que ela esteja.
 *
 * O cabeçalho corporativo não é feito de parágrafos soltos: é um grupo de
 * formas, e o título mora dentro de uma caixa. Ela vem inteira, porque digitar
 * dentro dela abre e fecha parágrafos — endereçar parágrafo a parágrafo
 * quebraria no primeiro Enter.
 */
export function editBandFloat(page: PageSetup, bid: string, content: DocumentNode[]): PageSetup {
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

/** Há algo a desenhar nesta faixa? */
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

/**
 * Aplica a mesma transformação às seis faixas que a configuração pode ter.
 *
 * São seis porque o Word distingue capa, páginas pares e o padrão, para
 * cabeçalho e para rodapé. Quem edita uma peça não sabe em qual delas ela está
 * — e não precisa saber: o endereço é único no arquivo, e trocar o texto atinge
 * todas as folhas de uma vez, como no Word. Escrever a lista das seis em cada
 * edição é como uma delas acaba esquecida numa.
 */
function mapBands(page: PageSetup, transform: (band: Band) => Band): PageSetup {
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

/**
 * Onde a faixa do cabeçalho e a do rodapé começam, medindo da borda da folha.
 *
 * Metade da margem: a faixa é mais larga que a coluna de texto, como o
 * documento corporativo a desenha — usar a margem do texto encolheria o
 * logotipo. Tela e papel leem daqui, e não cada um da sua conta: duas contas
 * iguais escritas em dois lugares é como os dois desenhos divergem.
 */
export function bandInsetMm(page: PageSetup): number {
  return Math.min(page.margins.left, page.margins.right) / 2
}
