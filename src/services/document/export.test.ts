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

  it('é uma página autocontida, com língua, título e sem script', () => {
    expect(output.startsWith('<!doctype html>')).toBe(true)
    expect(output).toContain('<html lang="pt-BR">')
    expect(output).toContain('<meta charset="utf-8">')
    expect(output).toContain('<title>Relatório anual</title>')
    expect(output).toContain('<meta name="author" content="Ana &quot;A&quot; Lima">')
    expect(output).not.toMatch(/<script/i)
    expect(output).toContain('Título &lt;script&gt;alert(1)&lt;/script&gt;')
  })

  it('cai no nome do arquivo quando as propriedades não têm título', () => {
    expect(html({ ...MODEL, properties: {} })).toContain('<title>relatorio</title>')
  })

  it('leva os estilos do documento e a formatação direta do parágrafo', () => {
    expect(output).toContain('.page__content > [data-style-id="Normal"]')
    expect(output).toContain('<p data-style-id="Normal" style="text-align: center">')
  })

  it('marca o texto com as tags semânticas', () => {
    expect(output).toContain('<strong>mundo</strong>')
    expect(output).toContain('<em>itálico</em>')
    expect(output).toContain('<s>riscado</s>')
    expect(output).toContain('<code>x*y</code>')
    expect(output).toContain('<a href="https://exemplo.com" rel="noopener noreferrer">site</a>')
    expect(output).toContain('42<br>linha 2 ')
    expect(output).toContain(`<img src="${PNG}" alt="logo">`)
  })

  it('exporta as revisões como texto final e deixa os comentários de fora', () => {
    expect(output).toContain(' novo')
    expect(output).not.toContain('velho')
    expect(output).not.toMatch(/<(ins|del)\b/)
    expect(output).toContain('<p>Junta isto</p>')
    expect(output).not.toContain('fora')
    expect(output).not.toContain('cid')
  })

  it('põe o marcador como âncora e o sumário como lista de links, sem número de página', () => {
    expect(output).toContain('<h1><a id="_Toc1"></a>Título')
    expect(output).toContain('<nav class="sumario"><ul>\n<li><a href="#_Toc1">Introdução</a></li>')
    expect(output).not.toContain('>7<')
  })

  it('numera as notas e as reúne no fim com o caminho de volta', () => {
    expect(output).toContain(
      '<sup class="nota-ref"><a href="#nota-rodape-1" id="ref-nota-rodape-1">1</a></sup>',
    )
    expect(output).toContain(
      '<li id="nota-rodape-1"><a href="#ref-nota-rodape-1" title="Voltar ao texto">1</a>',
    )
    expect(output).toContain('<p>Nota um</p>')
    // Endnotes count apart, in lowercase Roman numerals, as in Word.
    expect(output).toContain('href="#nota-fim-1" id="ref-nota-fim-1">i</a>')
    expect(output).toContain('<section class="notas" aria-label="Notas">')
  })

  it('mantém o início, a marca e o aninhamento das listas', () => {
    expect(output).toMatch(/<ol start="3" data-list-indent style="[^"]*">/)
    expect(output).toContain('data-label="3."')
    expect(output).toContain('data-label="4."')
    expect(output).toMatch(/<li[^>]*><p>quatro<\/p>\n<ul[^>]*>\n<li[^>]*><p>dentro<\/p><\/li>/)
  })

  it('exporta tabelas com larguras e células mescladas', () => {
    expect(output).toContain('<colgroup><col style="width: 100px"><col style="width: 200px"></colgroup>')
    expect(output).toContain('<td colspan="2"><p>mesclada</p></td>')
  })

  it('marca a quebra de página só para a impressão', () => {
    expect(output).toContain('<div class="page-break"></div>')
    expect(output).toContain('.page-break { display: block; break-after: page; }')
  })
})

describe('exportHtml — o que vem do documento não executa nem busca nada', () => {
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

  it('escapa texto, atributos e o título', () => {
    expect(output).not.toMatch(/<script|<img src=x|<svg|<b>/i)
    expect(output).toContain('&lt;img src=x onerror=alert(1)&gt;')
    expect(output).toContain('<title>&lt;/title&gt;&lt;script&gt;x&lt;/script&gt;</title>')
  })

  it('não leva link perigoso, CSS com url( nem imagem de fora', () => {
    expect(output).not.toContain('javascript:')
    expect(output.slice(output.indexOf('<body>'))).not.toMatch(/url\(|rastreio/)
    expect(output).toContain('clique')
  })

  it('bloqueia script e rede pela política da página', () => {
    expect(output).toContain("default-src 'none'; img-src data:")
  })

  it('aceita só o vocabulário de cores, medidas e fontes', () => {
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

  it('escreve títulos, ênfase, código e links', () => {
    expect(markdown).toContain('# <a id="_Toc1"></a>Título \\<script\\>alert(1)\\</script\\>')
    expect(markdown).toContain('Olá **mundo** *itálico* ~~riscado~~ `x*y` [site](https://exemplo.com) novo')
  })

  it('exporta o texto final das revisões', () => {
    expect(markdown).not.toContain('velho')
    expect(markdown).toContain('Junta isto')
    expect(markdown).not.toContain('fora')
  })

  it('quebra a linha com barra invertida e leva o resultado do campo', () => {
    expect(markdown).toContain('campo 42\\\nlinha 2')
  })

  it('grava a imagem na pasta irmã e a referencia por caminho relativo', () => {
    expect(markdown).toContain('![logo](relat%C3%B3rio_arquivos/image-1.png)')
    expect(assets).toEqual([{ name: 'image-1.png', mime: 'image/png', base64: 'iVBORw0KGgo=' }])
  })

  it('usa as notas do GFM para as de rodapé e as de fim', () => {
    expect(markdown).toContain('novo[^1]')
    expect(markdown).toContain('Fim[^fim-1]')
    expect(markdown).toContain('[^1]: Nota um')
    expect(markdown).toContain('[^fim-1]: Nota de fim')
  })

  it('mantém o início da lista numerada e o aninhamento', () => {
    expect(markdown).toContain('3. três\n4. quatro\n   - dentro')
  })

  it('escreve a tabela simples em GFM e a mesclada em HTML', () => {
    expect(markdown).toContain('| A | B\\|C |\n| --- | --- |\n| 1 | 2 |')
    expect(markdown).toContain('<table>')
    expect(markdown).toContain('<td colspan="2"><p>mesclada</p></td>')
  })

  it('escreve o sumário como lista de links, sem número de página', () => {
    expect(markdown).toContain('- [Introdução](#_Toc1)')
    expect(markdown).not.toContain('7')
  })

  it('escapa o que viraria marcação', () => {
    expect(markdown).toContain('\\# não é título')
    expect(markdown).toContain('1\\. não é lista')
    expect(markdown).toContain(
      '\\*estrela\\* \\_sub\\_ \\[x\\](y) \\<b\\>negrito\\</b\\> a\\|b \\~til\\~ \\\\ &amp;amp;',
    )
  })
})

describe('escapeMarkdown e as marcas', () => {
  it('escapa os caracteres especiais e a entidade', () => {
    expect(escapeMarkdown('a*b_c`d')).toBe('a\\*b\\_c\\`d')
    expect(escapeMarkdown('&copy; & co')).toBe('&amp;copy; & co')
  })

  it('tira o espaço das pontas da ênfase', () => {
    const doc: DocumentNode = {
      type: 'doc',
      content: [paragraph(text('a'), text(' forte ', [{ type: 'bold' }]), text('b'))],
    }
    expect(exportMarkdown({ doc }, { assetFolder: 'x' }).markdown).toBe('a **forte** b\n')
  })

  it('não abre ênfase só com espaço e não deixa delimitador vazio', () => {
    const doc: DocumentNode = {
      type: 'doc',
      content: [paragraph(text('a'), text(' ', [{ type: 'italic' }]), text('b'))],
    }
    expect(exportMarkdown({ doc }, { assetFolder: 'x' }).markdown).toBe('a b\n')
  })

  it('não leva link perigoso', () => {
    const doc: DocumentNode = {
      type: 'doc',
      content: [paragraph(text('x', [{ type: 'link', attrs: { href: 'javascript:alert(1)' } }]))],
    }
    expect(exportMarkdown({ doc }, { assetFolder: 'x' }).markdown).toBe('x\n')
  })

  it('cerca o código com mais crases do que ele tem', () => {
    const doc: DocumentNode = { type: 'doc', content: [paragraph(text('a`b', [{ type: 'code' }]))] }
    expect(exportMarkdown({ doc }, { assetFolder: 'x' }).markdown).toBe('``a`b``\n')
  })
})

describe('finalDocument', () => {
  it('não deixa célula vazia quando tudo nela foi excluído', () => {
    const doc = finalDocument({
      type: 'tableCell',
      content: [{ type: 'paragraph', content: [text('x', [{ type: 'deletion' }])] }],
    })
    expect(doc.content).toEqual([{ type: 'paragraph', content: [] }])
  })
})

describe('as equações (M11, fase 3)', () => {
  const FRACTION = '<math display="block"><mfrac><mi>a</mi><mi>b</mi></mfrac></math>'
  const inline = (attrs: Record<string, unknown>): DocumentNode => ({
    type: 'math',
    attrs: { omml: null, display: false, ...attrs },
  })
  const doc = (...content: DocumentNode[]): DocumentNode => ({ type: 'doc', content })

  it('no HTML, o MathML filtrado, em bloco só a de exibição', () => {
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

  it('no Markdown, o LaTeX entre cifrões; o da exibição na linha dele', () => {
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

  it('no Markdown, a equação sem LaTeX nenhum vai como MathML', () => {
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
