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
 * The editor delivers the document **already split into pages**, and each becomes a paper-sized
 * box: with `@page { margin: 0 }`, Chromium only stacks them. Letting it paginate would mean two
 * paginators that must agree. Bands are DOM inside the page, not Chromium's `headerTemplate`, which
 * draws the same band on every page.
 */

export interface PrintPage {
  readonly number: number
  readonly html: string
  readonly floats: readonly PrintFloat[]
  /** With `pageNumberStart` set to the section's first sheet number; see `sheetSetups`. */
  readonly setup: PageSetup
  /** The sheet within the section, from 1: it decides the title page and the number. */
  readonly inSection: number
  /** The blank sheet an even or odd section asked for. */
  readonly blank?: boolean
  /** This sheet's column separator lines (`w:cols/@w:sep`), in sheet mm. */
  readonly columnLines?: readonly {
    readonly leftMm: number
    readonly topMm: number
    readonly heightMm: number
  }[]
  readonly notes?: readonly PrintNoteArea[]
}

/** The same as the screen's (`NoteArea`), in mm and with the notes' HTML. */
export interface PrintNoteArea {
  readonly topMm: number
  readonly leftMm: number
  readonly widthMm: number
  readonly separator: 'normal' | 'continuation' | null
  readonly separatorMm: number
  readonly items: readonly {
    readonly html: string
    /** Where, in the note body, this sheet's first line starts. */
    readonly clipTopMm: number
    readonly heightMm: number
  }[]
}

/** The box text comes as serialized HTML: the editor holds the ProseMirror schema. */
export interface PrintFloat {
  readonly object: FloatingObject
  readonly anchorTopMm: number
  readonly contentHtml?: string | undefined
}

/**
 * `@page` is generated because size and orientation come from the document. The margin is zero: the
 * page box indents the text, and `printToPDF` would count it twice.
 */
export function buildPagedCss(pages: readonly Pick<PrintPage, 'setup'>[]): string {
  // One named `@page` per paper: a landscape sheet comes out landscape in the middle of a portrait
  // document. Chromium honors the name with `preferCSSPageSize`.
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
/* A extensão que o arquivo declara já conta o contorno. */
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

/* Topo e base vêm em linha, do que o documento declara. */
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
 * Band heights travel along because paper needs the same margin math the screen did: a header
 * taller than the margin pushes the body down.
 */
export interface PagedDocument {
  readonly pages: readonly PrintPage[]
  /** Per section, in section order (`useBandHeights`). */
  readonly bands: readonly BandHeights[]
  /** Each sheet's section, to find its band heights. */
  readonly sections?: readonly number[]
}

/** The dimensions, which is what tells papers apart. */
function paperName(size: { width: number; height: number }): string {
  return `folha-${Math.round(size.width * 10)}x${Math.round(size.height * 10)}`
}

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

  // HTML order is stacking order: what sits behind (`behindDoc`) comes first, the text in the
  // middle, the front last.
  const floats = sheet.floats

  return (
    `<div class="paper-page paper-page--${paperName(pageDimensionsMm(page))}">` +
    renderFloats(floats, page, true) +
    (hasBandContent(header)
      ? renderBand(header, 'header', pageLabel(page, sheet.inSection), total, {
          inset,
          offset: page.headerDistanceMm,
        })
      : '') +
    body +
    (hasBandContent(footer)
      ? renderBand(footer, 'footer', pageLabel(page, sheet.inSection), total, {
          inset,
          offset: page.footerDistanceMm,
        })
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

/** A continuation lifts the note body to the line where the previous sheet stopped. */
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
      // Around the center, as Word rotates: the box is positioned unrotated and rotated afterwards.
      (box.rotation === 0 ? '' : `transform:rotate(${box.rotation}deg);`) +
      // The same frame as the screen, through the same function.
      Object.entries(frameOf(item.object))
        .map(([property, value]) => `${property}:${value};`)
        .join('')

    if (item.object.kind === 'image') {
      return `<img class="paper-float" alt="" style="${style}" src="${escapeHtml(item.object.src ?? '')}" />`
    }

    // The rule: a flat, wide shape with an outline and no content, the line under a corporate
    // header.
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
  { inset, offset }: { readonly inset: number; readonly offset: number },
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

/** With cells already resolved by the reader, like the screen. */
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

/** For the band's three thirds and for the grid cells. */
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
