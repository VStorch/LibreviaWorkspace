import { contentInsetsMm, pageDimensionsMm, type PageSetup } from './model.js'
import {
  NO_BANDS,
  bandForPage,
  pageLabel,
  pieceText,
  bandInsetMm,
  hasBandContent,
  linesOf,
  type Band,
  type BandCell,
  type BandHeights,
  type BandPiece,
} from './band.js'
import { frameOf, placeFloating, type FloatingObject } from './floating.js'

/**
 * O papel montado a partir das mesmas páginas que a tela desenha.
 *
 * Até aqui havia **dois paginadores que precisavam concordar**: o nosso, na
 * tela, e o do Chromium, na exportação. A regra de "não deixar título sozinho no
 * pé da página" estava escrita duas vezes — uma em JavaScript e outra em CSS —
 * sem nada que forçasse a sincronia, e os dois arquivos comentavam esse risco um
 * para o outro. Bastava uma divergir para o PDF quebrar noutro lugar.
 *
 * Agora o editor entrega o documento **já dividido em páginas**, e cada uma vira
 * uma caixa do tamanho exato do papel. O Chromium deixa de decidir onde cortar:
 * com `@page { margin: 0 }` e uma caixa por folha, ele só empilha o que
 * recebeu. Some o paginador duplicado, e com ele a categoria inteira de defeito.
 *
 * O que se ganha além disso: cabeçalho e rodapé passam a ser DOM de verdade
 * dentro da página, em vez do `headerTemplate` do Chromium. O template roda num
 * contexto isolado, com escala própria — daí os fatores 0,75 e 0,6 codificados
 * em `services/pdf/page-setup.ts` — e desenha a **mesma** faixa em todas as
 * páginas, o que tornava impossível uma capa com cabeçalho próprio.
 */

/** O corte no papel entre uma folha e a seguinte é dado; aqui só se empilha. */
export interface PrintPage {
  /** Número da folha, começando em 1. */
  readonly number: number
  /** Os blocos daquela folha, no HTML que o editor produziu. */
  readonly html: string
  /** Os objetos ancorados que caem nesta folha. */
  readonly floats: readonly PrintFloat[]
  /**
   * A seção da folha (M9): papel, margens e faixas dela, com `pageNumberStart`
   * no número da primeira folha da seção — ver `sheetSetups`.
   */
  readonly setup: PageSetup
  /** A folha dentro da seção, a partir de 1: é o que decide a capa e o número. */
  readonly inSection: number
  /** A folha em branco que a seção par ou ímpar pediu. */
  readonly blank?: boolean
  /** As linhas entre colunas desta folha (`w:cols/@w:sep`), em mm da folha. */
  readonly columnLines?: readonly {
    readonly leftMm: number
    readonly topMm: number
    readonly heightMm: number
  }[]
  /** As áreas de notas desta folha (M11), já com o HTML de cada nota e o recorte da tela. */
  readonly notes?: readonly PrintNoteArea[]
}

/** Uma área de notas no papel: a mesma da tela (`NoteArea`), em mm e com o HTML das notas. */
export interface PrintNoteArea {
  readonly topMm: number
  readonly leftMm: number
  readonly widthMm: number
  readonly separator: 'normal' | 'continuation' | null
  readonly separatorMm: number
  readonly items: readonly {
    readonly html: string
    /** Onde, no corpo da nota, começa a primeira linha desta folha. */
    readonly clipTopMm: number
    readonly heightMm: number
  }[]
}

/**
 * Um objeto ancorado, pronto para a conta de posição.
 *
 * O conteúdo da caixa de texto vem em HTML já serializado: quem tem o schema do
 * ProseMirror é o editor, e este módulo desenha o papel sem saber que ele
 * existe. A posição, essa é calculada aqui — pela mesma função que a tela usa,
 * que é o que garante que os dois desenhem no mesmo lugar.
 */
export interface PrintFloat {
  readonly object: FloatingObject
  readonly anchorTopMm: number
  readonly contentHtml?: string | undefined
}

/**
 * Regras que só existem no papel paginado por nós.
 *
 * `@page` precisa ser gerado: tamanho e orientação vêm do documento, e uma
 * regra fixa numa folha de estilo compartilhada valeria para o papel errado.
 * A margem é zero **de propósito** — quem recua o texto é a caixa da página, e
 * pedir margem também ao `printToPDF` a contaria duas vezes.
 */
export function buildPagedCss(pages: readonly Pick<PrintPage, 'setup'>[]): string {
  // Um `@page` nomeado por papel: a folha em paisagem sai em paisagem no meio de
  // um documento em retrato. O Chromium honra o nome com `preferCSSPageSize`.
  const papers = new Map<string, { width: number; height: number }>()
  for (const sheet of pages) {
    const size = pageDimensionsMm(sheet.setup)
    papers.set(paperName(size), size)
  }
  const first = pages[0] === undefined ? { width: 210, height: 297 } : pageDimensionsMm(pages[0].setup)
  const named = [...papers]
    .map(
      ([name, size]) =>
        `@page ${name} { size: ${size.width}mm ${size.height}mm; margin: 0; }\n` +
        `.paper-page--${name} { page: ${name}; width: ${size.width}mm; height: ${size.height}mm; }`,
    )
    .join('\n')

  return `
@page { size: ${first.width}mm ${first.height}mm; margin: 0; }

.paper-page {
  position: relative;
  box-sizing: border-box;
  width: ${first.width}mm;
  height: ${first.height}mm;
  /* Bloco mais alto que a folha transborda na tela; no papel não há para onde
     transbordar, e deixá-lo invadir a folha seguinte sobreporia texto a texto. */
  overflow: hidden;
  break-after: page;
}

/* Sem isto o Chromium fecha o documento com uma folha em branco. */
.paper-page:last-child { break-after: auto; }

.paper-page__body { height: 100%; box-sizing: border-box; position: relative; z-index: 1; }

.paper-floats { position: absolute; inset: 0; }
.paper-floats--behind { z-index: 0; }
.paper-floats--front { z-index: 2; }
/* A caixa inclui o contorno: a extensão que o arquivo declara já o conta, e
   somá-lo por fora esticaria a forma pela espessura do traço. Crase nenhuma
   aqui dentro: isto mora num template literal. */
.paper-float { position: absolute; object-fit: contain; box-sizing: border-box; }
.paper-float--text > * { margin: 0; }

.paper-page__band {
  position: absolute;
  display: grid;
  grid-template-columns: auto 1fr auto;
  align-items: center;
  /* Só entre as colunas: ver a mesma regra em styles.css. */
  column-gap: 8px;
  font-size: 9pt;
  color: #222222;
}

/* Topo e base vêm em linha, do que o documento declara. Crase nenhuma
   aqui dentro: isto mora num template literal. */
.paper-page__band--ruled { border-bottom: 1px solid #999999; padding-bottom: 2px; }
.paper-page__band img { object-fit: contain; }
/* O filete do cabeçalho: a forma tem altura zero, e o que se ve e o contorno. */
.paper-float--rule { border-top: 1px solid #000000; }
/* Mesma regra de styles.css: o br ocupa a largura toda para quebrar a linha
   dentro do flex, que é como cada paragrafo do arquivo vira uma linha. */
.paper-page__cell { display: flex; flex-direction: column; justify-content: center; min-width: 0; }
.paper-page__line { display: flex; align-items: center; gap: 6px; }
.paper-page__cell--center { align-items: center; }
.paper-page__cell--right { align-items: flex-end; }

/* A grade atravessa os três terços: ela é a moldura do cabeçalho, não uma peça
   a ser distribuída entre eles. */
.paper-page__grid {
  grid-column: 1 / -1;
  width: 100%;
  border-collapse: collapse;
  table-layout: fixed;
}
.paper-page__grid td { padding: 0 1.9mm; vertical-align: middle; overflow-wrap: break-word; }
.paper-page__grid img { max-width: 100%; height: auto; }

.paper-column-line { position: absolute; width: 0; border-left: 1px solid #000000; }

/* Por último, para vencer a medida padrão de .paper-page acima: cada folha com
   o papel da sua seção. */
${named}
`
}

/**
 * O documento já recortado em folhas, com o que a tela mediu.
 *
 * As alturas das faixas viajam junto porque o papel precisa da **mesma** conta
 * de margem que a tela fez: um cabeçalho mais alto que a margem de cima desce o
 * corpo, e se os dois medissem por conta própria a folha da tela e a do papel
 * começariam em alturas diferentes.
 */
export interface PagedDocument {
  readonly pages: readonly PrintPage[]
  /** Altura das faixas de cada seção, na ordem das seções (`useBandHeights`). */
  readonly bands: readonly BandHeights[]
  /** A seção de cada folha, para achar a altura das faixas dela. */
  readonly sections?: readonly number[]
}

/** O nome do `@page` de um papel: as medidas, que é o que o distingue. */
function paperName(size: { width: number; height: number }): string {
  return `folha-${Math.round(size.width * 10)}x${Math.round(size.height * 10)}`
}

/** As folhas, uma caixa cada. */
export function buildPagedBody(paged: PagedDocument): string {
  const total = paged.pages.length

  return paged.pages
    .map((sheet, index) => {
      const page = sheet.setup
      const bands = paged.bands[paged.sections?.[index] ?? 0] ?? paged.bands[0] ?? NO_BANDS
      return renderPage(sheet, page, total, bandInsetMm(page), contentInsetsMm(page, bands))
    })
    .join('\n')
}

function renderPage(
  sheet: PrintPage,
  page: PageSetup,
  total: number,
  inset: number,
  insets: { top: number; bottom: number },
): string {
  const header = bandForPage(page, sheet.inSection, 'header')
  const footer = bandForPage(page, sheet.inSection, 'footer')

  const body =
    `<div class="page__content paper-page__body" style="padding:${insets.top}mm ${page.margins.right}mm ${insets.bottom}mm ${page.margins.left}mm">` +
    sheet.html +
    '</div>'

  // A ordem no HTML é a ordem de empilhamento, junto com o `z-index`: o que fica
  // atrás vem antes, o texto no meio, o que fica na frente por último. É a
  // distinção que o `behindDoc` do OOXML faz para decoração de capa.
  // Os objetos das faixas já vêm dentro de `sheet.floats`, com o texto das
  // caixas serializado: quem conhece o schema do ProseMirror é o editor.
  const floats = sheet.floats

  return (
    `<div class="paper-page paper-page--${paperName(pageDimensionsMm(page))}">` +
    renderFloats(floats, page, true) +
    (hasBandContent(header)
      ? renderBand(header, 'header', pageLabel(page, sheet.inSection), total, inset, page.headerDistanceMm)
      : '') +
    body +
    (hasBandContent(footer)
      ? renderBand(footer, 'footer', pageLabel(page, sheet.inSection), total, inset, page.footerDistanceMm)
      : '') +
    renderFloats(floats, page, false) +
    (sheet.columnLines ?? [])
      .map(
        (line) =>
          `<div class="paper-column-line" style="left:${line.leftMm}mm;top:${line.topMm}mm;height:${line.heightMm}mm"></div>`,
      )
      .join('') +
    (sheet.notes ?? []).map(renderNotes).join('') +
    '</div>'
  )
}

/**
 * A área de notas, desenhada como na tela: o separador e cada nota recortada na
 * altura que a paginação lhe deu — a continuação sobe o corpo até a linha em
 * que a folha anterior parou.
 */
function renderNotes(area: PrintNoteArea): string {
  const separator =
    area.separator === null
      ? ''
      : `<div class="paper-notes__separator${area.separator === 'continuation' ? ' paper-notes__separator--continued' : ''}" style="height:${area.separatorMm}mm"></div>`
  const items = area.items
    .map(
      (item) =>
        `<div class="paper-notes__slot" style="height:${item.heightMm}mm">` +
        `<div class="page__content note-body" style="margin-top:${-item.clipTopMm}mm">${item.html}</div></div>`,
    )
    .join('')
  return (
    `<div class="paper-notes" style="top:${area.topMm}mm;left:${area.leftMm}mm;width:${area.widthMm}mm">` +
    separator +
    items +
    '</div>'
  )
}

function renderFloats(floats: readonly PrintFloat[], page: PageSetup, behind: boolean): string {
  const visible = floats.filter((item) => item.object.behind === behind)
  if (visible.length === 0) return ''

  const boxes = visible.map((item) => {
    const box = placeFloating(item.object, page, item.anchorTopMm)
    const style =
      `left:${box.leftMm}mm;top:${box.topMm}mm;` +
      `width:${box.widthMm}mm;height:${box.heightMm}mm;` +
      // Em torno do centro, como o Word gira: a caixa é posicionada sem girar e
      // o giro acontece depois.
      (box.rotation === 0 ? '' : `transform:rotate(${box.rotation}deg);`) +
      // A mesma moldura da tela, pela mesma função: duas regras iguais escritas
      // em dois lugares é como os dois desenhos divergem. As chaves já são os
      // nomes das propriedades de CSS — `background` e `border`.
      Object.entries(frameOf(item.object))
        .map(([property, value]) => `${property}:${value};`)
        .join('')

    if (item.object.kind === 'image') {
      return `<img class="paper-float" alt="" style="${style}" src="${escapeHtml(item.object.src ?? '')}" />`
    }

    // O filete: forma rasa e larga, com contorno e sem conteúdo — a linha que
    // corre sob o cabeçalho corporativo.
    if (item.object.kind === 'rule') {
      return `<div class="paper-float paper-float--rule" style="${style}"></div>`
    }

    return `<div class="paper-float paper-float--text page__content" style="${style}">${item.contentHtml ?? ''}</div>`
  })

  return `<div class="paper-floats paper-floats--${behind ? 'behind' : 'front'}">${boxes.join('')}</div>`
}

function renderBand(
  band: Band,
  kind: 'header' | 'footer',
  label: string,
  total: number,
  inset: number,
  offset: number,
): string {
  const cell = (pieces: readonly BandPiece[], place: string): string =>
    `<div class="paper-page__cell paper-page__cell--${place}">` + renderLines(pieces, label, total) + '</div>'

  return (
    `<div class="paper-page__band paper-page__band--${kind}${band.rule ? ' paper-page__band--ruled' : ''}" ` +
    `style="left:${inset}mm;right:${inset}mm;${kind === 'header' ? 'top' : 'bottom'}:${offset}mm">` +
    renderGrid(band, label, total) +
    cell(band.left, 'left') +
    cell(band.center, 'center') +
    cell(band.right, 'right') +
    '</div>'
  )
}

/**
 * A grade do cabeçalho no papel.
 *
 * A mesma tabela que a tela desenha, a partir das mesmas células já resolvidas:
 * larguras, mesclagem e bordas vêm prontas do leitor, e nenhum dos dois refaz a
 * conta por conta própria — que é como tela e papel divergem.
 */
function renderGrid(band: Band, label: string, total: number): string {
  if (band.rows.length === 0) return ''

  const rows = band.rows
    .map((row) => {
      const cells = row.cells
        .map((cell) => {
          const span = cell.span === 1 ? '' : ` colspan="${cell.span}"`
          const down = cell.rowSpan === 1 ? '' : ` rowspan="${cell.rowSpan}"`
          const pieces = renderLines(cell.pieces, label, total)
          return `<td${span}${down} style="${cellStyle(cell)}">${pieces}</td>`
        })
        .join('')
      return `<tr>${cells}</tr>`
    })
    .join('')

  return `<table class="paper-page__grid"><tbody>${rows}</tbody></table>`
}

/**
 * As peças distribuídas em linhas, como o arquivo as quebrou.
 *
 * Vale para os três terços da faixa e para as células da grade: nos dois a
 * quebra vem do mesmo `linesOf`, e uma segunda versão desta marcação divergiria
 * da primeira no primeiro ajuste de estilo.
 */
function renderLines(pieces: readonly BandPiece[], label: string, total: number): string {
  return linesOf(pieces)
    .map(
      (line) =>
        '<div class="paper-page__line">' +
        line.map((piece) => renderPiece(piece, label, total)).join('') +
        '</div>',
    )
    .join('')
}

function cellStyle(cell: BandCell): string {
  const line = '1px solid currentcolor'
  const side = (initial: string, name: string): string =>
    cell.borders.includes(initial) ? `border-${name}:${line};` : ''

  return (
    (cell.width > 0 ? `width:${(cell.width * 100).toFixed(2)}%;` : '') +
    (cell.align === undefined ? '' : `text-align:${escapeHtml(cell.align)};`) +
    side('t', 'top') +
    side('l', 'left') +
    side('b', 'bottom') +
    side('r', 'right')
  )
}

function renderPiece(piece: BandPiece, label: string, total: number): string {
  if (piece.kind === 'image') {
    if (piece.src === undefined) return ''
    // Sem o fator de escala que o template do Chromium exigia: aqui a imagem
    // está numa página de verdade e a medida do documento vale como está.
    const width = piece.width === undefined ? '' : `width:${piece.width}px;`
    return `<img src="${escapeHtml(piece.src)}" alt="" style="${width}" />`
  }

  const text = pieceText(piece, label, total)

  const style =
    (piece.bold ? 'font-weight:700;' : '') +
    (piece.italic ? 'font-style:italic;' : '') +
    (piece.color === undefined ? '' : `color:${escapeHtml(piece.color)};`) +
    (piece.fontSize === undefined ? '' : `font-size:${escapeHtml(piece.fontSize)};`) +
    (piece.fontFamily === undefined ? '' : `font-family:${escapeHtml(piece.fontFamily)};`)

  return `<span style="${style}">${escapeHtml(text)}</span>`
}

function escapeHtml(text: string): string {
  return text
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
}
