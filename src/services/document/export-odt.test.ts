import { describe, expect, it } from 'vitest'
import { exportOdt, formulaSizeMm, odtEntries, pixelSizeOf } from './export-odt.js'
import { DEFAULT_PAGE_SETUP, type DocumentModel, type DocumentNode } from './model.js'
import { odfText, xml } from './odt-xml.js'
import { BUILTIN_STYLES } from './styles.js'
import { crc32, zip } from './zip.js'

/** Um PNG de 2 × 3 pixels, só o cabeçalho — o bastante para medir. */
const PNG = (() => {
  const bytes = new Uint8Array(33)
  bytes.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52])
  bytes.set([0, 0, 0, 2, 0, 0, 0, 3], 16)
  return `data:image/png;base64,${btoa(String.fromCharCode(...bytes))}`
})()

const text = (value: string, marks: DocumentNode['marks'] = undefined): DocumentNode =>
  marks === undefined ? { type: 'text', text: value } : { type: 'text', text: value, marks }
const paragraph = (...content: DocumentNode[]): DocumentNode => ({ type: 'paragraph', content })
const item = (...content: DocumentNode[]): DocumentNode => ({ type: 'listItem', content })
const cell = (value: string, attrs: Record<string, unknown> = {}): DocumentNode => ({
  type: 'tableCell',
  attrs: { colspan: 1, rowspan: 1, ...attrs },
  content: [paragraph(text(value))],
})

/** Um documento com tudo o que a exportação trata — também o do teste de ponta a ponta. */
export const RICH_ODT_MODEL: DocumentModel = {
  page: { ...DEFAULT_PAGE_SETUP, header: 'Relatório — página {n} de {total}' },
  sections: [
    {
      ...DEFAULT_PAGE_SETUP,
      id: 's1',
      margins: { top: 30, right: 20, bottom: 20, left: 30 },
    },
  ],
  styles: BUILTIN_STYLES,
  properties: {
    title: 'Relatório <anual> & "final"',
    subject: 'Contas',
    creator: 'Ana',
    keywords: 'contas; balanço',
    created: '2026-01-02T03:04:05Z',
    modified: '2026-02-03T04:05:06Z',
  },
  comments: [
    {
      id: '9',
      author: 'Bia <revisora>',
      date: '2026-03-01T10:00:00Z',
      paragraphs: ['Confira & ajuste'],
      done: false,
    },
    { id: '10', parentId: '9', author: 'Ana', date: '', paragraphs: ['Feito'], done: false },
  ],
  doc: {
    type: 'doc',
    content: [
      {
        type: 'heading',
        attrs: { level: 1 },
        content: [
          { type: 'bookmarkStart', attrs: { name: '_Toc1', bid: '1' } },
          text('Título <script>&'),
          { type: 'bookmarkEnd', attrs: { bid: '1' } },
        ],
      },
      {
        type: 'tableOfContents',
        attrs: { instr: 'TOC \\o "1-3" \\h', head: 0 },
        content: [
          {
            type: 'paragraph',
            attrs: { styleId: 'TOC1' },
            content: [
              text('Introdução', [{ type: 'link', attrs: { href: '#_Toc1' } }]),
              text('\t'),
              { type: 'field', attrs: { instr: 'PAGEREF _Toc1 \\h', result: '1' } },
            ],
          },
        ],
      },
      {
        type: 'paragraph',
        attrs: { textAlign: 'center', spaceBefore: 6, spaceAfter: 12, lineHeight: '1.5', keepNext: true },
        content: [
          text('Olá  '),
          text('negrito', [{ type: 'bold' }]),
          text(' '),
          text('itálico', [{ type: 'italic' }]),
          text(' sub', [{ type: 'underline' }, { type: 'strike' }]),
          text(' cor', [
            {
              type: 'textStyle',
              attrs: { color: '#c00000', fontFamily: 'Arial, sans-serif', fontSize: '14pt' },
            },
          ]),
          text(' realce', [{ type: 'highlight', attrs: { color: 'yellow' } }]),
          text(' link', [{ type: 'link', attrs: { href: 'https://example.com/?a=1&b=2' } }]),
          text(' perigoso', [{ type: 'link', attrs: { href: 'javascript:alert(1)' } }]),
          { type: 'noteRef', attrs: { kind: 'footnote' }, content: [paragraph(text('Nota de rodapé'))] },
          { type: 'commentStart', attrs: { cid: '9' } },
          text(' comentado'),
          { type: 'commentEnd', attrs: { cid: '9' } },
          text(' removido', [{ type: 'deletion' }]),
          text(' inserido', [{ type: 'insertion' }]),
          { type: 'hardBreak' },
          { type: 'field', attrs: { instr: 'PAGE', result: '1' } },
          { type: 'image', attrs: { src: PNG, alt: 'figura', width: 96, height: 48 } },
        ],
      },
      {
        type: 'orderedList',
        attrs: {
          numbering: {
            key: 'k1',
            levels: [
              { fmt: 'decimal', text: '%1.', start: 3 },
              { fmt: 'lowerLetter', text: '%1.%2)', start: 1 },
            ],
          },
        },
        content: [
          item(paragraph(text('três'))),
          item(paragraph(text('quatro')), { type: 'orderedList', content: [item(paragraph(text('sub')))] }),
        ],
      },
      { type: 'bulletList', content: [item(paragraph(text('marcador')))] },
      { type: 'pageBreak' },
      {
        type: 'table',
        content: [
          {
            type: 'tableRow',
            content: [
              cell('A', { colspan: 2, colwidth: [96, 96] }),
              cell('B', { rowspan: 2, colwidth: [96] }),
            ],
          },
          {
            type: 'tableRow',
            content: [
              {
                type: 'tableCell',
                attrs: { colspan: 1, rowspan: 1 },
                content: [
                  paragraph(text('C'), {
                    type: 'noteRef',
                    attrs: { kind: 'footnote' },
                    content: [paragraph(text('Nota na célula'))],
                  }),
                ],
              },
              cell('D', { borders: 'top:double,1.5,#ff0000', shading: '#eeeeee' }),
            ],
          },
        ],
      },
      {
        type: 'paragraph',
        attrs: { sectionBreak: 's1' },
        content: [
          text('Fim da primeira seção'),
          { type: 'noteRef', attrs: { kind: 'endnote' }, content: [paragraph(text('Nota de fim'))] },
        ],
      },
      paragraph(text('Segunda seção')),
    ],
  },
}

/** As entradas de um ZIP, lidas pelo diretório central — os dados como estão no arquivo. */
function unzip(bytes: Uint8Array): Map<string, { data: Uint8Array; method: number; extra: number }> {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength)
  let end = bytes.length - 22
  while (end > 0 && view.getUint32(end, true) !== 0x06054b50) end--
  const count = view.getUint16(end + 10, true)
  let at = view.getUint32(end + 16, true)
  const entries = new Map<string, { data: Uint8Array; method: number; extra: number }>()
  for (let index = 0; index < count; index++) {
    const method = view.getUint16(at + 10, true)
    const size = view.getUint32(at + 20, true)
    const nameLength = view.getUint16(at + 28, true)
    const extraLength = view.getUint16(at + 30, true)
    const commentLength = view.getUint16(at + 32, true)
    const offset = view.getUint32(at + 42, true)
    const name = decode(bytes.subarray(at + 46, at + 46 + nameLength))
    const localName = view.getUint16(offset + 26, true)
    const localExtra = view.getUint16(offset + 28, true)
    const start = offset + 30 + localName + localExtra
    entries.set(name, { data: bytes.subarray(start, start + size), method, extra: localExtra })
    at += 46 + nameLength + extraLength + commentLength
  }
  return entries
}

const decode = (bytes: Uint8Array | undefined): string => new TextDecoder().decode(bytes)

/** Um "deflate" de mentira, que só encolhe: o que se testa é o contêiner, não a compressão. */
const shrink = (data: Uint8Array): Uint8Array => data.subarray(0, 1)

describe('zip', () => {
  it('calcula o CRC-32 de referência', () => {
    expect(crc32(new TextEncoder().encode('123456789'))).toBe(0xcbf43926)
  })

  it('comprime o que encolhe, guarda o resto, e o CRC é o do original', () => {
    const big = new TextEncoder().encode('a'.repeat(1000))
    const entries = unzip(
      zip(
        [
          { name: 'x.txt', data: big },
          { name: 'ç.bin', data: new Uint8Array([1]) },
        ],
        shrink,
      ),
    )
    expect(entries.get('x.txt')).toMatchObject({ method: 8, data: big.subarray(0, 1) })
    expect(entries.get('ç.bin')?.method).toBe(0)
    expect([...(entries.get('ç.bin')?.data ?? [])]).toEqual([1])
  })
})

describe('exportOdt', () => {
  // Guardado, para que o teste leia o XML sem descomprimir.
  const bytes = exportOdt(RICH_ODT_MODEL)
  const entries = unzip(bytes)
  const read = (name: string): string => decode(entries.get(name)?.data)
  const content = read('content.xml')
  const styles = read('styles.xml')

  it('abre com o mimetype, primeiro e guardado, sem campo extra', () => {
    expect([...bytes.subarray(0, 4)]).toEqual([0x50, 0x4b, 0x03, 0x04])
    expect(decode(bytes.subarray(30, 38))).toBe('mimetype')
    expect(decode(bytes.subarray(38, 38 + 39))).toBe('application/vnd.oasis.opendocument.text')
    expect([...entries.keys()][0]).toBe('mimetype')
    expect(entries.get('mimetype')).toMatchObject({ method: 0, extra: 0 })
  })

  it('declara no manifesto cada parte e cada imagem', () => {
    expect([...entries.keys()]).toEqual([
      'mimetype',
      'content.xml',
      'styles.xml',
      'meta.xml',
      'Pictures/image1.png',
      'META-INF/manifest.xml',
    ])
    const manifest = read('META-INF/manifest.xml')
    expect(manifest).toContain(
      'manifest:full-path="/" manifest:version="1.3" manifest:media-type="application/vnd.oasis.opendocument.text"',
    )
    expect(manifest).toContain('manifest:full-path="Pictures/image1.png" manifest:media-type="image/png"')
    expect(manifest).toContain('manifest:full-path="content.xml" manifest:media-type="text/xml"')
  })

  it('escapa o texto, os atributos e os metadados', () => {
    expect(content).toContain('Título &lt;script&gt;&amp;')
    expect(content).not.toContain('<script>')
    expect(content).toContain('xlink:href="https://example.com/?a=1&amp;b=2"')
    expect(content).not.toContain('javascript:')
    expect(content).toContain('<dc:creator>Bia &lt;revisora&gt;</dc:creator>')
    const meta = read('meta.xml')
    expect(meta).toContain('<dc:title>Relatório &lt;anual&gt; &amp; &quot;final&quot;</dc:title>')
    expect(meta).toContain('<meta:keyword>contas</meta:keyword><meta:keyword>balanço</meta:keyword>')
    expect(meta).toContain('<meta:initial-creator>Ana</meta:initial-creator>')
    expect(meta).toContain('<meta:creation-date>2026-01-02T03:04:05Z</meta:creation-date>')
  })

  it('leva títulos, marcadores, links, notas, campos e comentários', () => {
    expect(content).toMatch(
      /<text:h text:style-name="[^"]+" text:outline-level="1"><text:bookmark-start text:name="_Toc1"\/>/,
    )
    expect(content).toContain('<text:bookmark-end text:name="_Toc1"/>')
    expect(content).toContain('<text:a xlink:type="simple" xlink:href="#_Toc1">')
    expect(content).toMatch(
      /<text:note text:id="nota-rodape-1" text:note-class="footnote"><text:note-citation>1<\/text:note-citation><text:note-body><text:p[^>]*>Nota de rodapé<\/text:p>/,
    )
    expect(content).toContain('text:note-class="endnote"><text:note-citation>i</text:note-citation>')
    expect(content).toContain('<text:page-number text:select-page="current">1</text:page-number>')
    expect(content).toContain('<office:annotation office:name="__Annotation__9">')
    expect(content).toContain('<office:annotation-end office:name="__Annotation__9"/>')
    expect(content).toContain('<text:p>Feito</text:p>')
    // A revisão sai aceita.
    expect(content).not.toContain('removido')
    expect(content).toContain('<text:s/>inserido')
  })

  it('guarda os espaços em sequência e a tabulação', () => {
    expect(content).toContain('Olá <text:s/>')
    expect(odfText('  a\tb   c')).toBe('<text:s text:c="2"/>a<text:tab/>b <text:s text:c="2"/>c')
    expect(xml('a\u0001<')).toBe('a&lt;')
  })

  it('formata o trecho com estilo automático e o parágrafo com a direta', () => {
    expect(styles).toMatch(
      /<style:style style:name="[^"]+" style:display-name="heading 1" style:family="paragraph"/,
    )
    expect(content).toMatch(
      /<style:style style:name="T\d+" style:family="text"><style:text-properties fo:font-weight="bold"/,
    )
    expect(content).toContain('fo:color="#c00000"')
    expect(content).toContain('style:font-name="Arial"')
    expect(content).toContain('fo:background-color="#ffff00"')
    expect(content).toContain('style:text-underline-style="solid"')
    expect(content).toContain('style:text-line-through-style="solid"')
    expect(content).toMatch(/fo:text-align="center"[^>]*fo:margin-top="6pt" fo:margin-bottom="12pt"/)
    expect(content).toContain('fo:keep-with-next="always"')
  })

  it('numera as listas como a conta do Word, aninhadas e com início', () => {
    expect(content).toMatch(/<text:list text:style-name="L1"><text:list-item text:start-value="3">/)
    expect(content).toContain(
      '<text:list-level-style-number text:level="1" style:num-suffix="." style:num-format="1" text:start-value="3">',
    )
    expect(content).toContain(
      'text:level="2" style:num-suffix=")" style:num-format="a" text:display-levels="2"',
    )
    expect(content).toContain('<text:list-level-style-bullet text:level="1" text:bullet-char="•">')
  })

  it('mescla as células com cobertas e põe a imagem no pacote', () => {
    expect(content).toContain('table:number-columns-spanned="2"')
    expect(content).toContain('table:number-rows-spanned="2"')
    // A nota de dentro da célula continua nota.
    expect(content).toMatch(/<table:table-cell[^>]*><text:p[^>]*>C<text:note [^>]*>.*Nota na célula/)
    // Linha 1: A (2 colunas) + coberta + B; linha 2: C, D e a coberta de B.
    expect(content.match(/<table:covered-table-cell\/>/g)).toHaveLength(2)
    expect(content).toContain('fo:border-top="1.5pt double #ff0000"')
    expect(content).toContain('<draw:image xlink:href="Pictures/image1.png"')
    expect(content).toContain('svg:width="25.4mm" svg:height="12.7mm"')
    const png = Uint8Array.from(atob(PNG.split(',')[1]!), (char) => char.charCodeAt(0))
    expect(pixelSizeOf(png)).toEqual({ width: 2, height: 3 })
  })

  it('dá a cada seção a página mestra, e a quebra ao bloco seguinte', () => {
    expect(content).toMatch(
      /style:master-page-name="Standard"><style:paragraph-properties[^>]*style:page-number="1"/,
    )
    expect(content).toMatch(/style:master-page-name="Section2"/)
    expect(content).toContain('fo:break-before="page"')
    expect(styles).toContain('<style:master-page style:name="Standard" style:page-layout-name="pm1">')
    expect(styles).toContain('fo:page-width="210mm" fo:page-height="297mm"')
    expect(styles).toContain('fo:margin-left="30mm"')
    // O cabeçalho de texto simples, centralizado por tabulação, com os campos de página.
    const header = /<style:header>(.*?)<\/style:header>/.exec(styles)?.[1] ?? ''
    expect(header).toMatch(/^<text:p text:style-name="MP\d+"><text:tab\/>/)
    expect(header).toContain('Relatório — página </text:span>')
    expect(header).toContain('<text:page-number text:select-page="current">1</text:page-number>')
    expect(header).toContain('<text:page-count>1</text:page-count>')
  })

  it('com compressão, só as partes de XML são comprimidas', () => {
    const packed = unzip(exportOdt(RICH_ODT_MODEL, {}, shrink))
    expect([...packed.entries()].filter(([, entry]) => entry.method === 8).map(([name]) => name)).toEqual([
      'content.xml',
      'styles.xml',
      'meta.xml',
      'META-INF/manifest.xml',
    ])
    expect(odtEntries(RICH_ODT_MODEL).map((entry) => entry.name)).toEqual([...entries.keys()])
  })
})

describe('as equações no ODT (M11, fase 3)', () => {
  const SQUARE = '<math display="inline"><msup><mi>x</mi><mn>2</mn></msup></math>'
  const FRACTION = '<math display="block"><mfrac><mi>a</mi><mi>b</mi></mfrac></math>'
  const model: DocumentModel = {
    ...RICH_ODT_MODEL,
    doc: {
      type: 'doc',
      content: [
        paragraph(
          text('Seja '),
          { type: 'math', attrs: { mathml: SQUARE, latex: 'x^2', display: false } },
          text(' & mais'),
        ),
        paragraph({
          type: 'math',
          attrs: { mathml: FRACTION, latex: '', display: true, omml: '<m:oMath/>' },
        }),
        paragraph({ type: 'math', attrs: { mathml: '<script/>', latex: 'y' } }),
      ],
    },
  }
  const entries = unzip(exportOdt(model))
  const read = (name: string): string => decode(entries.get(name)?.data)
  const content = read('content.xml')

  it('cada equação é um objeto de fórmula, como um caractere, com o MathML na pasta dele', () => {
    expect(content).toMatch(
      /<draw:frame draw:style-name="fr\d+" draw:name="Equation\d+" text:anchor-type="as-char" svg:width="[\d.]+mm" svg:height="[\d.]+mm" draw:z-index="\d+"><draw:object xlink:href="\.\/Object 1" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"\/><svg:desc>x\^2<\/svg:desc><\/draw:frame>/,
    )
    expect(content).toContain('xlink:href="./Object 2"')
    // A que não passa no filtro não deixa objeto nem texto.
    expect(content).not.toContain('Object 3')
    expect(content).not.toContain('script')
    expect(read('Object 1/content.xml')).toBe(
      '<?xml version="1.0" encoding="UTF-8"?>\n<math xmlns="http://www.w3.org/1998/Math/MathML" display="inline"><msup><mi>x</mi><mn>2</mn></msup></math>',
    )
    expect(read('Object 2/content.xml')).toContain(
      'display="block"><mfrac><mi>a</mi><mi>b</mi></mfrac></math>',
    )
    // O LaTeX que vai na descrição sai do MathML quando a equação não o guarda.
    expect(content).toContain('<svg:desc>\\frac{a}{b}</svg:desc>')
  })

  it('o manifesto declara cada objeto e cada parte dele, e só o que está no pacote', () => {
    const manifest = read('META-INF/manifest.xml')
    expect(manifest).toContain(
      '<manifest:file-entry manifest:full-path="Object 1/" manifest:version="1.3" manifest:media-type="application/vnd.oasis.opendocument.formula"/>',
    )
    expect(manifest).toContain('manifest:full-path="Object 2/content.xml" manifest:media-type="text/xml"')
    const declared = [...manifest.matchAll(/manifest:full-path="([^"]+)"/g)].map((match) => match[1]!)
    const files = declared.filter((path) => path !== '/' && !path.endsWith('/'))
    expect(files.sort()).toEqual(
      [...entries.keys()].filter((name) => name !== 'mimetype' && !name.startsWith('META-INF/')).sort(),
    )
    for (const folder of declared.filter((path) => path.length > 1 && path.endsWith('/'))) {
      expect(files.some((path) => path.startsWith(folder))).toBe(true)
    }
  })

  it('o quadro cresce com o texto e com o que a equação empilha', () => {
    const flat = formulaSizeMm({
      tag: 'math',
      attrs: {},
      children: [{ tag: 'mi', attrs: {}, children: ['x'] }],
    })
    const stacked = formulaSizeMm({
      tag: 'math',
      attrs: {},
      children: [
        {
          tag: 'mfrac',
          attrs: {},
          children: [
            { tag: 'mi', attrs: {}, children: ['x'] },
            { tag: 'mn', attrs: {}, children: ['2'] },
          ],
        },
      ],
    })
    expect(stacked.height).toBeGreaterThan(flat.height)
    expect(stacked.width).toBeGreaterThan(flat.width)
  })
})
