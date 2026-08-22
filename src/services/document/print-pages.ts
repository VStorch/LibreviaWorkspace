import {
  bandForPage,
  hasBandContent,
  pageDimensionsMm,
  type Band,
  type BandPiece,
  type PageSetup,
} from './model.js'

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
}

/**
 * Regras que só existem no papel paginado por nós.
 *
 * `@page` precisa ser gerado: tamanho e orientação vêm do documento, e uma
 * regra fixa numa folha de estilo compartilhada valeria para o papel errado.
 * A margem é zero **de propósito** — quem recua o texto é a caixa da página, e
 * pedir margem também ao `printToPDF` a contaria duas vezes.
 */
export function buildPagedCss(page: PageSetup): string {
  const { width, height } = pageDimensionsMm(page)

  return `
@page { size: ${width}mm ${height}mm; margin: 0; }

.paper-page {
  position: relative;
  box-sizing: border-box;
  width: ${width}mm;
  height: ${height}mm;
  /* Bloco mais alto que a folha transborda na tela; no papel não há para onde
     transbordar, e deixá-lo invadir a folha seguinte sobreporia texto a texto. */
  overflow: hidden;
  break-after: page;
}

/* Sem isto o Chromium fecha o documento com uma folha em branco. */
.paper-page:last-child { break-after: auto; }

.paper-page__body { height: 100%; box-sizing: border-box; }

.paper-page__band {
  position: absolute;
  display: grid;
  grid-template-columns: auto 1fr auto;
  align-items: center;
  gap: 8px;
  font-size: 9pt;
  color: #222222;
}

.paper-page__band--header { top: 4mm; }
.paper-page__band--footer { bottom: 4mm; }
.paper-page__band--ruled { border-bottom: 1px solid #999999; padding-bottom: 2px; }
.paper-page__band img { object-fit: contain; }
.paper-page__cell { display: flex; align-items: center; gap: 6px; min-width: 0; }
.paper-page__cell--center { justify-content: center; }
.paper-page__cell--right { justify-content: flex-end; }
`
}

/** As folhas, uma caixa cada. */
export function buildPagedBody(pages: readonly PrintPage[], page: PageSetup): string {
  const total = pages.length
  // Mesma proporção da tela: a faixa do cabeçalho corporativo é mais larga que
  // a coluna de texto, e usar a margem do texto encolheria o logotipo.
  const inset = Math.min(page.margins.left, page.margins.right) / 2

  return pages.map((sheet) => renderPage(sheet, page, total, inset)).join('\n')
}

function renderPage(sheet: PrintPage, page: PageSetup, total: number, inset: number): string {
  const header = bandForPage(page, sheet.number, 'header')
  const footer = bandForPage(page, sheet.number, 'footer')

  const body =
    `<div class="page__content paper-page__body" style="padding:${page.margins.top}mm ${page.margins.right}mm ${page.margins.bottom}mm ${page.margins.left}mm">` +
    sheet.html +
    '</div>'

  return (
    '<div class="paper-page">' +
    (hasBandContent(header) ? renderBand(header, 'header', sheet.number, total, inset) : '') +
    body +
    (hasBandContent(footer) ? renderBand(footer, 'footer', sheet.number, total, inset) : '') +
    '</div>'
  )
}

function renderBand(
  band: Band,
  kind: 'header' | 'footer',
  pageNumber: number,
  total: number,
  inset: number,
): string {
  const cell = (pieces: readonly BandPiece[], place: string): string =>
    `<div class="paper-page__cell paper-page__cell--${place}">` +
    pieces.map((piece) => renderPiece(piece, pageNumber, total)).join('') +
    '</div>'

  return (
    `<div class="paper-page__band paper-page__band--${kind}${band.rule ? ' paper-page__band--ruled' : ''}" ` +
    `style="left:${inset}mm;right:${inset}mm">` +
    cell(band.left, 'left') +
    cell(band.center, 'center') +
    cell(band.right, 'right') +
    '</div>'
  )
}

function renderPiece(piece: BandPiece, pageNumber: number, total: number): string {
  if (piece.kind === 'image') {
    if (piece.src === undefined) return ''
    // Sem o fator de escala que o template do Chromium exigia: aqui a imagem
    // está numa página de verdade e a medida do documento vale como está.
    const width = piece.width === undefined ? '' : `width:${piece.width}px;`
    return `<img src="${escapeHtml(piece.src)}" alt="" style="${width}" />`
  }

  const text =
    piece.kind === 'pageNumber'
      ? String(pageNumber)
      : piece.kind === 'totalPages'
        ? String(total)
        : (piece.text ?? '')

  const style =
    (piece.bold ? 'font-weight:700;' : '') +
    (piece.italic ? 'font-style:italic;' : '') +
    (piece.color === undefined ? '' : `color:${escapeHtml(piece.color)};`) +
    (piece.fontSize === undefined ? '' : `font-size:${escapeHtml(piece.fontSize)};`)

  return `<span style="${style}">${escapeHtml(text)}</span>`
}

function escapeHtml(text: string): string {
  return text
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
}
