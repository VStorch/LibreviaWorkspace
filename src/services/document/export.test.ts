import { describe, expect, it } from 'vitest'
import { exportHtml, safeCss } from './export-html.js'
import { escapeMarkdown, exportMarkdown } from './export-markdown.js'
import { finalDocument } from './export-common.js'
import type { DocumentModel, DocumentNode } from './model.js'
import { BUILTIN_STYLES } from './styles.js'

const PNG = 'data:image/png;base64,iVBORw0KGgo='

const text = (value: string, marks: DocumentNode['marks'] = undefined): DocumentNode =>
  marks === undefined ? { type: 'text', text: value } : { type: 'text', text: value, marks }
const paragraph = (...content: DocumentNode[]): DocumentNode => ({ type: 'paragraph', content })
const cell = (value: string, attrs: Record<string, unknown> = {}): DocumentNode => ({
  type: 'tableCell',
  attrs: { colspan: 1, rowspan: 1, ...attrs },
  content: [paragraph(text(value))],
})
const item = (...content: DocumentNode[]): DocumentNode => ({ type: 'listItem', content })

/** Everything the export handles. */
const RICH: DocumentNode = {
  type: 'doc',
  content: [
    {
      type: 'heading',
      attrs: { level: 1 },
      content: [
        { type: 'bookmarkStart', attrs: { name: '_Toc1', bid: '1' } },
        text('Título <script>alert(1)</script>'),
        { type: 'bookmarkEnd', attrs: { bid: '1' } },
      ],
    },
    {
      type: 'tableOfContents',
      attrs: { instr: 'TOC \\o "1-3" \\h', head: 0, sdt: true },
      content: [
        {
          type: 'paragraph',
          attrs: { styleId: 'TOC1' },
          content: [
            text('Introdução', [{ type: 'link', attrs: { href: '#_Toc1' } }]),
            text('\t', [{ type: 'link', attrs: { href: '#_Toc1' } }]),
            { type: 'field', attrs: { instr: 'PAGEREF _Toc1 \\h', result: '7' } },
          ],
        },
      ],
    },
    {
      type: 'paragraph',
      attrs: { styleId: 'Normal', textAlign: 'center' },
      content: [
        text('Olá '),
        text('mundo', [{ type: 'bold' }]),
        text(' '),
        text('itálico', [{ type: 'italic' }]),
        text(' '),
        text('riscado', [{ type: 'strike' }]),
        text(' '),
        text('x*y', [{ type: 'code' }]),
        text(' '),
        text('site', [{ type: 'link', attrs: { href: 'https://exemplo.com' } }]),
        text(' novo', [{ type: 'insertion', attrs: { author: 'Ana' } }]),
        text(' velho', [{ type: 'deletion', attrs: { author: 'Ana' } }]),
        { type: 'noteRef', attrs: { kind: 'footnote' }, content: [paragraph(text('Nota um'))] },
        { type: 'commentStart', attrs: { cid: '9' } },
        text(' campo '),
        { type: 'commentEnd', attrs: { cid: '9' } },
        { type: 'field', attrs: { instr: 'SEQ Figura', result: '42' } },
        { type: 'hardBreak' },
        text('linha 2 '),
        { type: 'image', attrs: { src: PNG, alt: 'logo' } },
      ],
    },
    {
      type: 'paragraph',
      attrs: { markRevision: { kind: 'del', author: 'Ana' } },
      content: [text('Junta ')],
    },
    paragraph(text('isto')),
    {
      type: 'orderedList',
      attrs: { numbering: { key: 'k1', levels: [{ fmt: 'decimal', text: '%1.', start: 3 }] } },
      content: [
        item(paragraph(text('três'))),
        item(paragraph(text('quatro')), { type: 'bulletList', content: [item(paragraph(text('dentro')))] }),
      ],
    },
    {
      type: 'table',
      content: [
        { type: 'tableRow', content: [cell('A', { colwidth: [100] }), cell('B|C', { colwidth: [200] })] },
        { type: 'tableRow', content: [cell('1'), cell('2')] },
        { type: 'tableRow', attrs: { rowRevision: { kind: 'del' } }, content: [cell('fora'), cell('fora')] },
      ],
    },
    {
      type: 'table',
      content: [
        { type: 'tableRow', content: [cell('mesclada', { colspan: 2 })] },
        { type: 'tableRow', content: [cell('x'), cell('y')] },
      ],
    },
    paragraph(text('# não é título')),
    paragraph(text('1. não é lista')),
    paragraph(text('*estrela* _sub_ [x](y) <b>negrito</b> a|b ~til~ \\ &amp;')),
    { type: 'pageBreak' },
    paragraph(text('Fim'), {
      type: 'noteRef',
      attrs: { kind: 'endnote' },
      content: [paragraph(text('Nota de fim'))],
    }),
  ],
}

const MODEL: Pick<DocumentModel, 'doc' | 'styles' | 'notes' | 'properties'> = {
  doc: RICH,
  styles: BUILTIN_STYLES,
  properties: { title: 'Relatório anual', creator: 'Ana "A" Lima' },
}

const LABELS = { notes: 'Notas', backToText: 'Voltar ao texto' }
const html = (model = MODEL): string =>
  exportHtml(model, { fileName: 'relatorio.docx', lang: 'pt-BR', labels: LABELS })

describe('exportHtml', () => {
  const output = html()

  it('is a self-contained page, with language, title and no script', () => {
    expect(output.startsWith('<!doctype html>')).toBe(true)
    expect(output).toContain('<html lang="pt-BR">')
    expect(output).toContain('<meta charset="utf-8">')
    expect(output).toContain('<title>Relatório anual</title>')
    expect(output).toContain('<meta name="author" content="Ana &quot;A&quot; Lima">')
    expect(output).not.toMatch(/<script/i)
    expect(output).toContain('Título &lt;script&gt;alert(1)&lt;/script&gt;')
  })

  it('falls back to the file name when the properties have no title', () => {
    expect(html({ ...MODEL, properties: {} })).toContain('<title>relatorio</title>')
  })

  it("carries the document styles and the paragraph's direct formatting", () => {
    expect(output).toContain('.page__content > [data-style-id="Normal"]')
    expect(output).toContain('<p data-style-id="Normal" style="text-align: center">')
  })

  it('marks the text with semantic tags', () => {
    expect(output).toContain('<strong>mundo</strong>')
    expect(output).toContain('<em>itálico</em>')
    expect(output).toContain('<s>riscado</s>')
    expect(output).toContain('<code>x*y</code>')
    expect(output).toContain('<a href="https://exemplo.com" rel="noopener noreferrer">site</a>')
    expect(output).toContain('42<br>linha 2 ')
    expect(output).toContain(`<img src="${PNG}" alt="logo">`)
  })

  it('exports revisions as final text and leaves comments out', () => {
    expect(output).toContain(' novo')
    expect(output).not.toContain('velho')
    expect(output).not.toMatch(/<(ins|del)\b/)
    expect(output).toContain('<p>Junta isto</p>')
    expect(output).not.toContain('fora')
    expect(output).not.toContain('cid')
  })

  it('puts the bookmark as an anchor and the table of contents as a link list, without page numbers', () => {
    expect(output).toContain('<h1><a id="_Toc1"></a>Título')
    expect(output).toContain('<nav class="toc"><ul>\n<li><a href="#_Toc1">Introdução</a></li>')
    expect(output).not.toContain('>7<')
  })

  it('numbers the notes and gathers them at the end with a way back', () => {
    expect(output).toContain('<sup class="note-ref"><a href="#footnote-1" id="ref-footnote-1">1</a></sup>')
    expect(output).toContain('<li id="footnote-1"><a href="#ref-footnote-1" title="Voltar ao texto">1</a>')
    expect(output).toContain('<p>Nota um</p>')
    // Endnotes count apart, in lowercase Roman numerals, as in Word.
    expect(output).toContain('href="#endnote-1" id="ref-endnote-1">i</a>')
    expect(output).toContain('<section class="notes" aria-label="Notas">')
  })

  it('keeps the start, the marker and the nesting of lists', () => {
    expect(output).toMatch(/<ol start="3" data-list-indent style="[^"]*">/)
    expect(output).toContain('data-label="3."')
    expect(output).toContain('data-label="4."')
    expect(output).toMatch(/<li[^>]*><p>quatro<\/p>\n<ul[^>]*>\n<li[^>]*><p>dentro<\/p><\/li>/)
  })

  it('exports tables with widths and merged cells', () => {
    expect(output).toContain('<colgroup><col style="width: 100px"><col style="width: 200px"></colgroup>')
    expect(output).toContain('<td colspan="2"><p>mesclada</p></td>')
  })

  it('marks the page break for printing only', () => {
    expect(output).toContain('<div class="page-break"></div>')
    expect(output).toContain('.page-break { display: block; break-after: page; }')
  })
})

describe('exportHtml: what comes from the document neither runs nor fetches anything', () => {
  const hostile: DocumentNode = {
    type: 'doc',
    content: [
      {
        type: 'paragraph',
        attrs: { styleId: '"><script>x</script>', fontFamily: "x; background: url('http://rastreio')" },
        content: [
          text('clique', [{ type: 'link', attrs: { href: 'javascript:alert(1)' } }]),
          text('cor', [{ type: 'textStyle', attrs: { color: 'red;background:url(http://x)' } }]),
          text('<img src=x onerror=alert(1)>'),
          { type: 'image', attrs: { src: 'https://rastreio.example/pixel.png', alt: '"><b>' } },
          { type: 'bookmarkStart', attrs: { name: '"><svg onload=alert(1)>' } },
        ],
      },
    ],
  }
  const output = html({ ...MODEL, doc: hostile, properties: { title: '</title><script>x</script>' } })

  it('escapes text, attributes and the title', () => {
    expect(output).not.toMatch(/<script|<img src=x|<svg|<b>/i)
    expect(output).toContain('&lt;img src=x onerror=alert(1)&gt;')
    expect(output).toContain('<title>&lt;/title&gt;&lt;script&gt;x&lt;/script&gt;</title>')
  })

  it('carries no dangerous link, CSS with url( or outside image', () => {
    expect(output).not.toContain('javascript:')
    expect(output.slice(output.indexOf('<body>'))).not.toMatch(/url\(|rastreio/)
    expect(output).toContain('clique')
  })

  it('blocks script and network through the page policy', () => {
    expect(output).toContain("default-src 'none'; img-src data:")
  })

  it('accepts only the vocabulary of colors, measures and fonts', () => {
    expect(safeCss('#ff0000')).toBe('#ff0000')
    expect(safeCss("'Times New Roman', serif")).toBe("'Times New Roman', serif")
    expect(safeCss('12pt')).toBe('12pt')
    expect(safeCss('red; color: blue')).toBeNull()
    expect(safeCss('url(http://x)')).toBeNull()
    expect(safeCss('expression(alert(1))')).toBeNull()
    expect(safeCss('a\\62 c')).toBeNull()
  })
})

describe('exportMarkdown', () => {
  const { markdown, assets } = exportMarkdown(MODEL, { assetFolder: 'relatório_arquivos' })

  it('writes headings, emphasis, code and links', () => {
    expect(markdown).toContain('# <a id="_Toc1"></a>Título \\<script\\>alert(1)\\</script\\>')
    expect(markdown).toContain('Olá **mundo** *itálico* ~~riscado~~ `x*y` [site](https://exemplo.com) novo')
  })

  it('exports the final text of revisions', () => {
    expect(markdown).not.toContain('velho')
    expect(markdown).toContain('Junta isto')
    expect(markdown).not.toContain('fora')
  })

  it('breaks the line with a backslash and carries the field result', () => {
    expect(markdown).toContain('campo 42\\\nlinha 2')
  })

  it('saves the image in the sibling folder and references it by relative path', () => {
    expect(markdown).toContain('![logo](relat%C3%B3rio_arquivos/image-1.png)')
    expect(assets).toEqual([{ name: 'image-1.png', mime: 'image/png', base64: 'iVBORw0KGgo=' }])
  })

  it('uses GFM notes for footnotes and endnotes', () => {
    expect(markdown).toContain('novo[^1]')
    expect(markdown).toContain('Fim[^endnote-1]')
    expect(markdown).toContain('[^1]: Nota um')
    expect(markdown).toContain('[^endnote-1]: Nota de fim')
  })

  it('keeps the start of the numbered list and the nesting', () => {
    expect(markdown).toContain('3. três\n4. quatro\n   - dentro')
  })

  it('writes the simple table in GFM and the merged one in HTML', () => {
    expect(markdown).toContain('| A | B\\|C |\n| --- | --- |\n| 1 | 2 |')
    expect(markdown).toContain('<table>')
    expect(markdown).toContain('<td colspan="2"><p>mesclada</p></td>')
  })

  it('writes the table of contents as a link list, without page numbers', () => {
    expect(markdown).toContain('- [Introdução](#_Toc1)')
    expect(markdown).not.toContain('7')
  })

  it('escapes what would turn into markup', () => {
    expect(markdown).toContain('\\# não é título')
    expect(markdown).toContain('1\\. não é lista')
    expect(markdown).toContain(
      '\\*estrela\\* \\_sub\\_ \\[x\\](y) \\<b\\>negrito\\</b\\> a\\|b \\~til\\~ \\\\ &amp;amp;',
    )
  })
})

describe('escapeMarkdown and the marks', () => {
  it('escapes special characters and the entity', () => {
    expect(escapeMarkdown('a*b_c`d')).toBe('a\\*b\\_c\\`d')
    expect(escapeMarkdown('&copy; & co')).toBe('&amp;copy; & co')
  })

  it('trims the space at the ends of the emphasis', () => {
    const doc: DocumentNode = {
      type: 'doc',
      content: [paragraph(text('a'), text(' forte ', [{ type: 'bold' }]), text('b'))],
    }
    expect(exportMarkdown({ doc }, { assetFolder: 'x' }).markdown).toBe('a **forte** b\n')
  })

  it('does not open emphasis with only space and leaves no empty delimiter', () => {
    const doc: DocumentNode = {
      type: 'doc',
      content: [paragraph(text('a'), text(' ', [{ type: 'italic' }]), text('b'))],
    }
    expect(exportMarkdown({ doc }, { assetFolder: 'x' }).markdown).toBe('a b\n')
  })

  it('carries no dangerous link', () => {
    const doc: DocumentNode = {
      type: 'doc',
      content: [paragraph(text('x', [{ type: 'link', attrs: { href: 'javascript:alert(1)' } }]))],
    }
    expect(exportMarkdown({ doc }, { assetFolder: 'x' }).markdown).toBe('x\n')
  })

  it('fences the code with more backticks than it contains', () => {
    const doc: DocumentNode = { type: 'doc', content: [paragraph(text('a`b', [{ type: 'code' }]))] }
    expect(exportMarkdown({ doc }, { assetFolder: 'x' }).markdown).toBe('``a`b``\n')
  })
})

describe('finalDocument', () => {
  it('leaves no empty cell when everything in it was deleted', () => {
    const doc = finalDocument({
      type: 'tableCell',
      content: [{ type: 'paragraph', content: [text('x', [{ type: 'deletion' }])] }],
    })
    expect(doc.content).toEqual([{ type: 'paragraph', content: [] }])
  })
})

describe('equations (M11, phase 3)', () => {
  const FRACTION = '<math display="block"><mfrac><mi>a</mi><mi>b</mi></mfrac></math>'
  const inline = (attrs: Record<string, unknown>): DocumentNode => ({
    type: 'math',
    attrs: { omml: null, display: false, ...attrs },
  })
  const doc = (...content: DocumentNode[]): DocumentNode => ({ type: 'doc', content })

  it('in HTML, the filtered MathML, as a block only the display one', () => {
    const output = html({
      doc: doc(
        paragraph(text('Seja '), inline({ mathml: '<math display="block"><mi>x</mi></math>', latex: 'x' })),
        paragraph({ type: 'math', attrs: { mathml: FRACTION, latex: '\\frac{a}{b}', display: true } }),
      ),
      styles: BUILTIN_STYLES,
    })
    expect(output).toContain(
      '<math xmlns="http://www.w3.org/1998/Math/MathML" display="inline"><mi>x</mi></math>',
    )
    expect(output).toContain('display="block"><mfrac><mi>a</mi><mi>b</mi></mfrac></math>')
  })

  it('in Markdown, the LaTeX between dollar signs; the display one on its own line', () => {
    const markdown = exportMarkdown(
      {
        doc: doc(
          paragraph(
            text('Custa $5 e '),
            inline({ mathml: '<math><mi>x</mi></math>', latex: 'x^2' }),
            text('.'),
          ),
          paragraph(
            text('Logo'),
            { type: 'math', attrs: { mathml: FRACTION, latex: '', display: true } },
            text('fim'),
          ),
        ),
      },
      { assetFolder: 'x' },
    ).markdown
    // Without stored LaTeX (the equation came from a file), it comes from MathML.
    expect(markdown).toBe('Custa \\$5 e $x^2$.\n\nLogo\n$$\\frac{a}{b}$$\nfim\n')
  })

  it('in Markdown, an equation with no LaTeX at all goes as MathML', () => {
    const markdown = exportMarkdown(
      { doc: doc(paragraph(inline({ mathml: '<math><mi>x</mi></math>', latex: '', omml: '<m:oMath/>' }))) },
      { assetFolder: 'x' },
    ).markdown
    expect(markdown).toBe('$x$\n')
    const empty = exportMarkdown(
      { doc: doc(paragraph(text('a'), inline({ mathml: '<math></math>', latex: '' }))) },
      { assetFolder: 'x' },
    ).markdown
    expect(empty).toBe('a<math xmlns="http://www.w3.org/1998/Math/MathML" display="inline"></math>\n')
  })
})
