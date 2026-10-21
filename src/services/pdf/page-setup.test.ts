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
  it('converts margins from millimeters to inches', () => {
    // Chromium works in inches; a mistake here produces a PDF with silently wrong margins.
    const options = buildPrintOptions(
      withPage({ margins: { top: 25.4, right: 12.7, bottom: 25.4, left: 50.8 } }),
    )

    expect(options.margins.top).toBeCloseTo(1, 6)
    expect(options.margins.right).toBeCloseTo(0.5, 6)
    expect(options.margins.left).toBeCloseTo(2, 6)
  })

  it('translates size and orientation', () => {
    expect(buildPrintOptions(withPage({ size: PageSize.A4 })).pageSize).toBe('A4')
    expect(buildPrintOptions(withPage({ size: PageSize.Letter })).pageSize).toBe('Letter')
    expect(buildPrintOptions(withPage({ orientation: PageOrientation.Landscape })).landscape).toBe(true)
    expect(buildPrintOptions(withPage({ orientation: PageOrientation.Portrait })).landscape).toBe(false)
  })

  it('sempre imprime fundos', () => {
    // Otherwise text highlight and table header backgrounds vanish from the PDF.
    expect(buildPrintOptions(DEFAULT_PAGE_SETUP).printBackground).toBe(true)
  })

  it('does not turn on header and footer when there is none', () => {
    expect(buildPrintOptions(DEFAULT_PAGE_SETUP).displayHeaderFooter).toBe(false)
  })

  it.each([
    ['header only', { header: 'Relatório', footer: '' }],
    ['footer only', { header: '', footer: 'página {n}' }],
    ['both', { header: 'Relatório', footer: 'página {n}' }],
  ])('turns on header and footer when there is %s', (_label, overrides) => {
    expect(buildPrintOptions(withPage(overrides)).displayHeaderFooter).toBe(true)
  })

  it('ignores text that is only whitespace', () => {
    expect(buildPrintOptions(withPage({ header: '   ' })).displayHeaderFooter).toBe(false)
  })

  it('lets the page size come from the options, not from the CSS', () => {
    // Two sources of truth for the margin would double it.
    expect(buildPrintOptions(DEFAULT_PAGE_SETUP).preferCSSPageSize).toBe(false)
  })
})

describe('buildNativePrintOptions', () => {
  it('uses pixels for margins, not inches', () => {
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

  it('marks the margins as custom', () => {
    expect(buildNativePrintOptions(DEFAULT_PAGE_SETUP).margins.marginType).toBe('custom')
  })

  it('carries size, orientation and background to the native dialog', () => {
    const options = buildNativePrintOptions(
      withPage({ size: PageSize.Letter, orientation: PageOrientation.Landscape }),
    )

    expect(options.pageSize).toBe('Letter')
    expect(options.landscape).toBe(true)
    expect(options.printBackground).toBe(true)
  })
})

describe('buildHeaderFooterTemplate', () => {
  it('replaces {n} with the page number', () => {
    expect(buildHeaderFooterTemplate('Página {n}')).toContain('<span class="pageNumber"></span>')
  })

  it('replaces {total} with the total number of pages', () => {
    expect(buildHeaderFooterTemplate('{n} de {total}')).toContain('<span class="totalPages"></span>')
  })

  it('declara a fonte explicitamente', () => {
    // Without a declared font size, Chromium renders the template at size zero and the header
    // simply does not appear.
    expect(buildHeaderFooterTemplate('Relatório')).toMatch(/font-size:\s*9pt/)
  })

  it('escapes markup typed by the user', () => {
    const template = buildHeaderFooterTemplate('<script>alert(1)</script>')

    expect(template).not.toContain('<script>')
    expect(template).toContain('&lt;script&gt;')
  })

  it('escapes before inserting the placeholders, without breaking them', () => {
    const template = buildHeaderFooterTemplate('a & b — página {n}')

    expect(template).toContain('&amp;')
    expect(template).toContain('<span class="pageNumber"></span>')
  })

  it('returns an empty element when there is no text', () => {
    expect(buildHeaderFooterTemplate('')).toBe('<span></span>')
    expect(buildHeaderFooterTemplate('   ')).toBe('<span></span>')
  })
})

describe('preserved document band', () => {
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

  it('takes over the display when present, ignoring the typed text', () => {
    // The file's header has a logo and numbering; the text field would represent none of that.
    const options = buildPrintOptions(withPage({ header: 'texto qualquer', headerBand: band }))

    expect(options.headerTemplate).toContain('RELATÓRIO INTERNO')
    expect(options.headerTemplate).not.toContain('texto qualquer')
  })

  it('embeds the image, because the template does not fetch external resources', () => {
    // A logo by URL would simply not appear in the PDF.
    expect(buildBandTemplate(band)).toContain('src="data:image/png;base64,AAAA"')
  })

  it('keeps the placeholders that Chromium replaces', () => {
    expect(buildBandTemplate(band)).toContain('<span class="pageNumber"></span>')
  })

  it('draws the rule when the document has one', () => {
    expect(buildBandTemplate(band)).toMatch(/border-bottom:\s*1px solid/)
    expect(buildBandTemplate({ ...band, rule: false })).not.toMatch(/border-bottom/)
  })

  it('turns on header and footer even without typed text', () => {
    expect(buildPrintOptions(withPage({ header: '', headerBand: band })).displayHeaderFooter).toBe(true)
  })

  it('ignores an empty band and falls back to the typed text', () => {
    const empty = { left: [], center: [], right: [], rule: false, floats: [], rows: [] }
    const options = buildPrintOptions(withPage({ header: 'Relatório', headerBand: empty }))

    expect(options.headerTemplate).toContain('Relatório')
  })

  it('escapes markup coming from the document', () => {
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
