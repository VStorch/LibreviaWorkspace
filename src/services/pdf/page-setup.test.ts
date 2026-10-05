import { describe, expect, it } from 'vitest'
import { DEFAULT_PAGE_SETUP, PageOrientation, PageSize, type PageSetup } from '@services/document/model.js'
import {
  MIN_MARGIN_FOR_HEADER_MM,
  buildBandTemplate,
  buildHeaderFooterTemplate,
  buildNativePrintOptions,
  buildPrintOptions,
  marginFitsHeaderOrFooter,
} from './page-setup.js'

const withPage = (overrides: Partial<PageSetup>): PageSetup => ({ ...DEFAULT_PAGE_SETUP, ...overrides })

describe('buildPrintOptions', () => {
  it('converte as margens de milímetros para polegadas', () => {
    // Chromium works in inches; a mistake here produces a PDF with silently wrong margins.
    const options = buildPrintOptions(
      withPage({ margins: { top: 25.4, right: 12.7, bottom: 25.4, left: 50.8 } }),
    )

    expect(options.margins.top).toBeCloseTo(1, 6)
    expect(options.margins.right).toBeCloseTo(0.5, 6)
    expect(options.margins.left).toBeCloseTo(2, 6)
  })

  it('traduz tamanho e orientação', () => {
    expect(buildPrintOptions(withPage({ size: PageSize.A4 })).pageSize).toBe('A4')
    expect(buildPrintOptions(withPage({ size: PageSize.Letter })).pageSize).toBe('Letter')
    expect(buildPrintOptions(withPage({ orientation: PageOrientation.Landscape })).landscape).toBe(true)
    expect(buildPrintOptions(withPage({ orientation: PageOrientation.Portrait })).landscape).toBe(false)
  })

  it('sempre imprime fundos', () => {
    // Otherwise text highlight and table header backgrounds vanish from the PDF.
    expect(buildPrintOptions(DEFAULT_PAGE_SETUP).printBackground).toBe(true)
  })

  it('não liga cabeçalho e rodapé quando não há nenhum', () => {
    expect(buildPrintOptions(DEFAULT_PAGE_SETUP).displayHeaderFooter).toBe(false)
  })

  it.each([
    ['só cabeçalho', { header: 'Relatório', footer: '' }],
    ['só rodapé', { header: '', footer: 'página {n}' }],
    ['ambos', { header: 'Relatório', footer: 'página {n}' }],
  ])('liga cabeçalho e rodapé quando há %s', (_label, overrides) => {
    expect(buildPrintOptions(withPage(overrides)).displayHeaderFooter).toBe(true)
  })

  it('ignora texto que é só espaço em branco', () => {
    expect(buildPrintOptions(withPage({ header: '   ' })).displayHeaderFooter).toBe(false)
  })

  it('deixa o tamanho de página vir das opções, não do CSS', () => {
    // Two sources of truth for the margin would double it.
    expect(buildPrintOptions(DEFAULT_PAGE_SETUP).preferCSSPageSize).toBe(false)
  })
})

describe('buildNativePrintOptions', () => {
  it('usa pixels nas margens, não polegadas', () => {
    // The trap: printToPDF wants inches and print() wants pixels. Swapping them gets the margin
    // wrong by a factor of 96.
    const options = buildNativePrintOptions(
      withPage({ margins: { top: 25.4, right: 25.4, bottom: 25.4, left: 25.4 } }),
    )

    expect(options.margins.top).toBe(96)
    expect(
      buildPrintOptions(withPage({ margins: { top: 25.4, right: 25.4, bottom: 25.4, left: 25.4 } })).margins
        .top,
    ).toBeCloseTo(1, 6)
  })

  it('marca as margens como personalizadas', () => {
    expect(buildNativePrintOptions(DEFAULT_PAGE_SETUP).margins.marginType).toBe('custom')
  })

  it('leva tamanho, orientação e fundo para o diálogo nativo', () => {
    const options = buildNativePrintOptions(
      withPage({ size: PageSize.Letter, orientation: PageOrientation.Landscape }),
    )

    expect(options.pageSize).toBe('Letter')
    expect(options.landscape).toBe(true)
    expect(options.printBackground).toBe(true)
  })
})

describe('buildHeaderFooterTemplate', () => {
  it('substitui {n} pelo número da página', () => {
    expect(buildHeaderFooterTemplate('Página {n}')).toContain('<span class="pageNumber"></span>')
  })

  it('substitui {total} pelo total de páginas', () => {
    expect(buildHeaderFooterTemplate('{n} de {total}')).toContain('<span class="totalPages"></span>')
  })

  it('declara a fonte explicitamente', () => {
    // Without a declared font size, Chromium renders the template at size zero and the header
    // simply does not appear.
    expect(buildHeaderFooterTemplate('Relatório')).toMatch(/font-size:\s*9pt/)
  })

  it('escapa a marcação digitada pelo usuário', () => {
    const template = buildHeaderFooterTemplate('<script>alert(1)</script>')

    expect(template).not.toContain('<script>')
    expect(template).toContain('&lt;script&gt;')
  })

  it('escapa antes de inserir os marcadores, sem quebrá-los', () => {
    const template = buildHeaderFooterTemplate('a & b — página {n}')

    expect(template).toContain('&amp;')
    expect(template).toContain('<span class="pageNumber"></span>')
  })

  it('devolve um elemento vazio quando não há texto', () => {
    expect(buildHeaderFooterTemplate('')).toBe('<span></span>')
    expect(buildHeaderFooterTemplate('   ')).toBe('<span></span>')
  })
})

describe('faixa preservada do documento', () => {
  const band = {
    left: [{ kind: 'pageNumber' as const, bold: false, italic: false }],
    center: [
      { kind: 'text' as const, text: 'RELATÓRIO INTERNO', bold: true, italic: false, fontSize: '20pt' },
    ],
    right: [
      {
        kind: 'image' as const,
        src: 'data:image/png;base64,AAAA',
        width: 182,
        height: 40,
        bold: false,
        italic: false,
      },
    ],
    rule: true,
    floats: [],
    rows: [],
  }

  it('manda na exibição quando existe, ignorando o texto digitado', () => {
    // The file's header has a logo and numbering; the text field would represent none of that.
    const options = buildPrintOptions(withPage({ header: 'texto qualquer', headerBand: band }))

    expect(options.headerTemplate).toContain('RELATÓRIO INTERNO')
    expect(options.headerTemplate).not.toContain('texto qualquer')
  })

  it('embute a imagem, porque o template não busca recurso externo', () => {
    // A logo by URL would simply not appear in the PDF.
    expect(buildBandTemplate(band)).toContain('src="data:image/png;base64,AAAA"')
  })

  it('mantém os marcadores que o Chromium substitui', () => {
    expect(buildBandTemplate(band)).toContain('<span class="pageNumber"></span>')
  })

  it('desenha o filete quando o documento tem', () => {
    expect(buildBandTemplate(band)).toMatch(/border-bottom:\s*1px solid/)
    expect(buildBandTemplate({ ...band, rule: false })).not.toMatch(/border-bottom/)
  })

  it('liga cabeçalho e rodapé mesmo sem texto digitado', () => {
    expect(buildPrintOptions(withPage({ header: '', headerBand: band })).displayHeaderFooter).toBe(true)
  })

  it('ignora faixa vazia e volta para o texto digitado', () => {
    const empty = { left: [], center: [], right: [], rule: false, floats: [], rows: [] }
    const options = buildPrintOptions(withPage({ header: 'Relatório', headerBand: empty }))

    expect(options.headerTemplate).toContain('Relatório')
  })

  it('escapa marcação vinda do documento', () => {
    const hostile = {
      ...band,
      center: [{ kind: 'text' as const, text: '<script>alert(1)</script>', bold: false, italic: false }],
    }

    expect(buildBandTemplate(hostile)).not.toContain('<script>')
    expect(buildBandTemplate(hostile)).toContain('&lt;script&gt;')
  })
})

describe('marginFitsHeaderOrFooter', () => {
  it('recusa margem apertada demais', () => {
    // Chromium draws header and footer inside the margin and clips the excess: with a small margin,
    // the text vanishes without warning.
    expect(marginFitsHeaderOrFooter(MIN_MARGIN_FOR_HEADER_MM - 1)).toBe(false)
    expect(marginFitsHeaderOrFooter(MIN_MARGIN_FOR_HEADER_MM)).toBe(true)
    expect(marginFitsHeaderOrFooter(25)).toBe(true)
  })
})
