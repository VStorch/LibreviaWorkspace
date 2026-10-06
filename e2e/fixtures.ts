/**
 * `.docx` documents built in code, readable in review, like the sidecar fixtures. The ZIP is
 * written by hand with stored entries, uncompressed: only header, data and index, no dependency.
 */

import { Buffer } from 'node:buffer'

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
</Types>`

const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`

const DOCUMENT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
</Relationships>`

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
const WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'

function paragraph(text: string): string {
  return `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
}

/** The paragraph anchoring the comment: it is what makes the loss possible. */
const COMMENTED_PARAGRAPH =
  `<w:p><w:commentRangeStart w:id="1"/>` +
  `<w:r><w:t xml:space="preserve">Segundo parágrafo, com um comentário ancorado.</w:t></w:r>` +
  `<w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p>`

function documentXml(body: string): string {
  // Drawing namespaces always go in, so no new fixture is born without them.
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W}" xmlns:mc="${MC}" xmlns:wp="${WP}" xmlns:a="${A}" xmlns:wps="${WPS}" mc:Ignorable="wps"><w:body>${body}<w:sectPr/></w:body></w:document>`
}

const COMMENTS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="${W}">
<w:comment w:id="1" w:author="Revisor" w:date="2026-01-01T00:00:00Z"><w:p><w:r><w:t>Conferir este número.</w:t></w:r></w:p></w:comment>
</w:comments>`

/** The table opening the locked document; see `docxWithComment`. */
const LEADING_TABLE = ((): string => {
  const cell = (text: string): string =>
    `<w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>${paragraph(text)}</w:tc>`
  return (
    `<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid>` +
    `<w:tr>${cell('Item')}${cell('Valor')}</w:tr>` +
    `<w:tr>${cell('Café')}${cell('12')}</w:tr></w:tbl>`
  )
})()

/**
 * A comment anchored on the second paragraph. With `leadingTable`, a table opens the document,
 * where the cursor is and Table menu commands act.
 */
export async function docxWithComment(options: { leadingTable?: boolean } = {}): Promise<Buffer> {
  const table = options.leadingTable === true ? LEADING_TABLE : ''

  return zip([
    ['[Content_Types].xml', CONTENT_TYPES],
    ['_rels/.rels', ROOT_RELS],
    ['word/_rels/document.xml.rels', DOCUMENT_RELS],
    [
      'word/document.xml',
      documentXml(table + paragraph('Ata da reunião de terça.') + COMMENTED_PARAGRAPH + paragraph('Fim.')),
    ],
    ['word/comments.xml', COMMENTS_XML],
  ])
}

/**
 * A document with a bit of every revision the editor reads: insertion and deletion in the second
 * paragraph, an inserted paragraph mark, a moved range and a deleted table row. With
 * `leadingTable`, the same opening table as `docxWithComment`.
 */
export async function docxWithTrackedChange(options: { leadingTable?: boolean } = {}): Promise<Buffer> {
  const table = options.leadingTable === true ? LEADING_TABLE : ''
  const date = 'w:date="2026-01-01T00:00:00Z"'
  const tracked =
    `<w:p><w:r><w:t xml:space="preserve">Segundo parágrafo, </w:t></w:r>` +
    `<w:ins w:id="1" w:author="Revisor" ${date}><w:r><w:t>com uma inserção revisada.</w:t></w:r></w:ins>` +
    `<w:del w:id="2" w:author="Revisora" ${date}><w:r><w:delText xml:space="preserve"> Trecho excluído.</w:delText></w:r></w:del></w:p>` +
    `<w:p><w:pPr><w:rPr><w:ins w:id="3" w:author="Revisor" ${date}/></w:rPr></w:pPr>` +
    `<w:r><w:t>Parágrafo partido</w:t></w:r></w:p>` +
    `<w:p><w:moveFromRangeStart w:id="4" w:author="Revisor" ${date} w:name="move1"/>` +
    `<w:moveFrom w:id="5" w:author="Revisor" ${date}><w:r><w:t>Frase movida.</w:t></w:r></w:moveFrom>` +
    `<w:moveFromRangeEnd w:id="4"/></w:p>` +
    `<w:p><w:r><w:t xml:space="preserve">Destino: </w:t></w:r>` +
    `<w:moveToRangeStart w:id="6" w:author="Revisor" ${date} w:name="move1"/>` +
    `<w:moveTo w:id="7" w:author="Revisor" ${date}><w:r><w:t>Frase movida.</w:t></w:r></w:moveTo>` +
    `<w:moveToRangeEnd w:id="6"/></w:p>` +
    `<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>` +
    `<w:tr><w:trPr><w:del w:id="8" w:author="Revisora" ${date}/></w:trPr><w:tc>${paragraph('Linha excluída')}</w:tc></w:tr>` +
    `<w:tr><w:tc>${paragraph('Linha que fica')}</w:tc></w:tr></w:tbl>`

  return zip([
    ['[Content_Types].xml', CONTENT_TYPES.replace(/<Override PartName="\/word\/comments[^>]+>/, '')],
    ['_rels/.rels', ROOT_RELS],
    [
      'word/document.xml',
      documentXml(table + paragraph('Ata da reunião de terça.') + tracked + paragraph('Fim.')),
    ],
  ])
}

/**
 * A cell inserted by revision (`w:cellIns`), which locks editing: the editor does not represent
 * structure revisions. With `leadingTable`, the table opens the document.
 */
export async function docxWithCellRevision(options: { leadingTable?: boolean } = {}): Promise<Buffer> {
  const revised = LEADING_TABLE.replace(
    '<w:tcW w:w="4500" w:type="dxa"/></w:tcPr>',
    '<w:tcW w:w="4500" w:type="dxa"/><w:cellIns w:id="1" w:author="Revisor" w:date="2026-01-01T00:00:00Z"/></w:tcPr>',
  )
  const body =
    options.leadingTable === true
      ? revised + paragraph('Ata da reunião de terça.') + paragraph('Fim.')
      : paragraph('Ata da reunião de terça.') + revised + paragraph('Fim.')

  return zip([
    ['[Content_Types].xml', CONTENT_TYPES.replace(/<Override PartName="\/word\/comments[^>]+>/, '')],
    ['_rels/.rels', ROOT_RELS],
    ['word/document.xml', documentXml(body)],
  ])
}

/**
 * A footnote in the second paragraph, which does not lock. With `leadingTable`, the opening table
 * of `docxWithComment`.
 */
export async function docxWithFootnote(options: { leadingTable?: boolean } = {}): Promise<Buffer> {
  const table = options.leadingTable === true ? LEADING_TABLE : ''
  const noted =
    `<w:p><w:r><w:t xml:space="preserve">Segundo parágrafo, com uma nota.</w:t></w:r>` +
    `<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r></w:p>`
  const footnotes =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:footnotes xmlns:w="${W}">` +
    `<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>` +
    `<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>` +
    `<w:footnote w:id="1"><w:p><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> Fonte: ata anterior.</w:t></w:r></w:p></w:footnote>` +
    `</w:footnotes>`

  return zip([
    [
      '[Content_Types].xml',
      CONTENT_TYPES.replace(
        /<Override PartName="\/word\/comments[^>]+>/,
        '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>',
      ),
    ],
    ['_rels/.rels', ROOT_RELS],
    [
      'word/_rels/document.xml.rels',
      DOCUMENT_RELS.replace(
        'relationships/comments" Target="comments.xml"',
        'relationships/footnotes" Target="footnotes.xml"',
      ),
    ],
    [
      'word/document.xml',
      documentXml(table + paragraph('Ata da reunião de terça.') + noted + paragraph('Fim.')),
    ],
    ['word/footnotes.xml', footnotes],
  ])
}

const M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'

/** An `m:r` as Word writes it, with the math font in `w:rPr`. */
function mathRun(text: string, sty?: string): string {
  const style = sty === undefined ? '' : `<m:rPr><m:sty m:val="${sty}"/></m:rPr>`
  return `<m:r>${style}<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr><m:t>${text}</m:t></m:r>`
}

/** The text around the inline equation of `docxWithEquations`. */
export const EQUATION_BEFORE = 'A área do círculo é '
export const EQUATION_AFTER = ' para todo raio.'

/**
 * A document with equations: an inline one (πr²) mid-sentence, a centered display one (the
 * quadratic formula) and a locked one, a box without its top side, which the screen does not draw.
 * The `m` namespace goes on each equation, as LibreOffice declares it.
 */
export async function docxWithEquations(): Promise<Buffer> {
  const inline =
    `<w:p><w:r><w:t xml:space="preserve">${EQUATION_BEFORE}</w:t></w:r>` +
    `<m:oMath xmlns:m="${M}">${mathRun('π')}<m:sSup><m:e>${mathRun('r')}</m:e><m:sup>${mathRun('2')}</m:sup></m:sSup></m:oMath>` +
    `<w:r><w:t xml:space="preserve">${EQUATION_AFTER}</w:t></w:r></w:p>`
  const display =
    `<w:p><m:oMathPara xmlns:m="${M}"><m:oMathParaPr><m:jc m:val="center"/></m:oMathParaPr><m:oMath>` +
    `${mathRun('x')}${mathRun('=')}<m:f><m:num>${mathRun('-b±')}<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/>` +
    `<m:e>${mathRun('Δ')}</m:e></m:rad></m:num><m:den>${mathRun('2a')}</m:den></m:f>` +
    `</m:oMath></m:oMathPara></w:p>`
  const locked =
    `<w:p><w:r><w:t xml:space="preserve">Caixa: </w:t></w:r><m:oMath xmlns:m="${M}">` +
    `<m:borderBox><m:borderBoxPr><m:hideTop m:val="1"/></m:borderBoxPr><m:e>${mathRun('z')}</m:e></m:borderBox>` +
    `</m:oMath></w:p>`

  return zip([
    ['[Content_Types].xml', CONTENT_TYPES.replace(/<Override PartName="\/word\/comments[^>]+>/, '')],
    ['_rels/.rels', ROOT_RELS],
    [
      'word/document.xml',
      documentXml(paragraph('Equações do relatório.') + inline + display + locked + paragraph('Fim.')),
    ],
  ])
}

/** The properties of `docxWithProperties`, as Word writes them. */
export const PROPERTIES_CORE =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
  `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">` +
  `<dc:title>Relatório anual</dc:title><dc:subject>Contas</dc:subject><dc:creator>Ana Lima</dc:creator>` +
  `<cp:keywords>contas; 2025</cp:keywords><cp:lastModifiedBy>Bia</cp:lastModifiedBy><cp:revision>7</cp:revision>` +
  `<dcterms:created xsi:type="dcterms:W3CDTF">2025-01-02T03:04:05Z</dcterms:created>` +
  `<dcterms:modified xsi:type="dcterms:W3CDTF">2025-02-03T04:05:06Z</dcterms:modified>` +
  `<cp:contentStatus>Rascunho</cp:contentStatus></cp:coreProperties>`

export const PROPERTIES_APP =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
  `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">` +
  `<Template>Normal.dotm</Template><TotalTime>42</TotalTime><Company>ACME</Company></Properties>`

export const PROPERTIES_CUSTOM =
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
  `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">` +
  `<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Cliente"><vt:lpwstr>XPTO</vt:lpwstr></property></Properties>`

/** A document with `docProps/core.xml`, `app.xml` and `custom.xml`. */
export async function docxWithProperties(): Promise<Buffer> {
  const docProps = (name: string, type: string): string =>
    `<Override PartName="/docProps/${name}.xml" ContentType="application/vnd.openxmlformats-${type}+xml"/>`
  return zip([
    [
      '[Content_Types].xml',
      CONTENT_TYPES.replace(
        /<Override PartName="\/word\/comments[^>]+>/,
        docProps('core', 'package.core-properties') +
          docProps('app', 'officedocument.extended-properties') +
          docProps('custom', 'officedocument.custom-properties'),
      ),
    ],
    [
      '_rels/.rels',
      ROOT_RELS.replace(
        '</Relationships>',
        `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>` +
          `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>` +
          `<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>` +
          '</Relationships>',
      ),
    ],
    ['word/document.xml', documentXml(paragraph('Relatório de contas do ano.') + paragraph('Fim.'))],
    ['docProps/core.xml', PROPERTIES_CORE],
    ['docProps/app.xml', PROPERTIES_APP],
    ['docProps/custom.xml', PROPERTIES_CUSTOM],
  ])
}

const W14 = 'http://schemas.microsoft.com/office/word/2010/wordml'
const W15 = 'http://schemas.microsoft.com/office/word/2012/wordml'

/**
 * A comment with a reply and a resolved comment, as Word writes them: the reply embraces the
 * parent's range, and `commentsExtended.xml` links both through `w14:paraId` and marks the resolved
 * one with `w15:done`.
 */
export async function docxWithCommentThread(): Promise<Buffer> {
  const comment = (id: string, author: string, paraId: string, text: string): string =>
    `<w:comment w:id="${id}" w:author="${author}" w:initials="${author[0]}" w:date="2026-03-02T10:00:00Z">` +
    `<w:p w14:paraId="${paraId}"><w:r><w:t>${text}</w:t></w:r></w:p></w:comment>`
  const comments =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<w:comments xmlns:w="${W}" xmlns:w14="${W14}">` +
    comment('0', 'Ana', '10000000', 'Conferir o valor.') +
    comment('1', 'Bruno', '10000001', 'Conferido na planilha.') +
    comment('2', 'Carla', '10000002', 'Trocar o título.') +
    `</w:comments>`
  const extended =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<w15:commentsEx xmlns:w15="${W15}" xmlns:mc="${MC}" mc:Ignorable="w15">` +
    `<w15:commentEx w15:paraId="10000000" w15:done="0"/>` +
    `<w15:commentEx w15:paraId="10000001" w15:paraIdParent="10000000" w15:done="0"/>` +
    `<w15:commentEx w15:paraId="10000002" w15:done="1"/>` +
    `</w15:commentsEx>`
  const reference = (id: string): string =>
    `<w:commentRangeEnd w:id="${id}"/><w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="${id}"/></w:r>`

  return zip([
    [
      '[Content_Types].xml',
      CONTENT_TYPES.replace(
        '</Types>',
        '<Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/></Types>',
      ),
    ],
    ['_rels/.rels', ROOT_RELS],
    [
      'word/_rels/document.xml.rels',
      DOCUMENT_RELS.replace(
        '</Relationships>',
        '<Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/></Relationships>',
      ),
    ],
    [
      'word/document.xml',
      documentXml(
        paragraph('Ata da reunião de terça.') +
          `<w:p><w:r><w:t xml:space="preserve">O orçamento é de </w:t></w:r>` +
          `<w:commentRangeStart w:id="0"/><w:commentRangeStart w:id="1"/>` +
          `<w:r><w:t>doze mil reais</w:t></w:r>${reference('0')}${reference('1')}` +
          `<w:r><w:t xml:space="preserve"> por ano.</w:t></w:r></w:p>` +
          `<w:p><w:commentRangeStart w:id="2"/><w:r><w:t>Título provisório</w:t></w:r>${reference('2')}</w:p>` +
          paragraph('Fim.'),
      ),
    ],
    ['word/comments.xml', comments],
    ['word/commentsExtended.xml', extended],
  ])
}

/**
 * An anchored text box, which the editor does not redraw: if it survives, the block's original XML
 * was preserved.
 */
export async function docxWithTextBox(): Promise<Buffer> {
  const caixa =
    `<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps">` +
    `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">` +
    `<wp:extent cx="2000000" cy="1000000"/><wp:docPr id="1" name="Caixa"/>` +
    `<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
    `<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr/>` +
    `<wps:txbx><w:txbxContent><w:p><w:r><w:t>Título na caixa</w:t></w:r></w:p></w:txbxContent></wps:txbx>` +
    `<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic>` +
    `</wp:inline></w:drawing></mc:Choice>` +
    `<mc:Fallback><w:p><w:r><w:t>Título na caixa</w:t></w:r></w:p></mc:Fallback>` +
    `</mc:AlternateContent></w:r></w:p>`

  return zip([
    ['[Content_Types].xml', CONTENT_TYPES.replace(/<Override PartName="\/word\/comments[^>]+>/, '')],
    ['_rels/.rels', ROOT_RELS],
    ['word/document.xml', documentXml(caixa + paragraph('Primeiro parágrafo do corpo.'))],
  ])
}

/** A 4 × 4 PNG, square so it differs from the 400 × 100 box the document asks for. */
const SQUARE_PNG = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAQAAAAECAIAAAAmkwkpAAAADklEQVR4nGNwQAIMxHEAOEMMAfoZu1cAAAAASUVORK5CYII=',
  'base64',
)

const IMAGE_CONTENT_TYPES = CONTENT_TYPES.replace(
  '<Default Extension="xml"',
  '<Default Extension="png" ContentType="image/png"/><Default Extension="xml"',
).replace(/<Override PartName="\/word\/comments[^>]+>/, '')

const IMAGE_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/quadrado.png"/>
</Relationships>`

/**
 * An image stretched in the flow: `wp:extent` gives the size on the page, without the file's ratio.
 * The square PNG is declared as 400 × 100 px, 3810000 × 952500 EMU.
 */
export async function docxWithStretchedImage(description?: string): Promise<Buffer> {
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
  const PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
  const imagem =
    `<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">` +
    `<wp:extent cx="3810000" cy="952500"/>` +
    `<wp:docPr id="1" name="Quadrado"${description === undefined ? '' : ` descr="${description}"`}/>` +
    `<a:graphic><a:graphicData uri="${PIC}"><pic:pic xmlns:pic="${PIC}">` +
    `<pic:nvPicPr><pic:cNvPr id="1" name="Quadrado"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip xmlns:r="${R}" r:embed="rId9"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3810000" cy="952500"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
    `</pic:pic></a:graphicData></a:graphic>` +
    `</wp:inline></w:drawing></w:r></w:p>`

  return zip([
    ['[Content_Types].xml', IMAGE_CONTENT_TYPES],
    ['_rels/.rels', ROOT_RELS],
    ['word/_rels/document.xml.rels', IMAGE_RELS],
    ['word/media/quadrado.png', SQUARE_PNG],
    ['word/document.xml', documentXml(imagem + paragraph('Legenda da imagem.'))],
  ])
}

/** The same image, with alt text (`wp:docPr/@descr`) that goes through the schema. */
export async function docxWithDescribedImage(): Promise<Buffer> {
  return docxWithStretchedImage('Quadrado azul de teste')
}

/**
 * A table with what the editor represents: cell shading and border, horizontal merge and header
 * row, each an attribute or node type the schema may return differently.
 */
export async function docxWithStyledCells(): Promise<Buffer> {
  const bordas =
    `<w:tcBorders><w:top w:val="double" w:sz="12" w:color="FF0000"/>` +
    `<w:bottom w:val="nil"/></w:tcBorders>`
  const cell = (text: string, extra = ''): string =>
    `<w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/>${extra}</w:tcPr>${paragraph(text)}</w:tc>`

  const table =
    `<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>` +
    `<w:tr><w:trPr><w:tblHeader/></w:trPr>` +
    `${cell('Cabeçalho sombreado', '<w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/>')}` +
    `${cell('Cabeçalho B')}${cell('Cabeçalho C')}</w:tr>` +
    `<w:tr>${cell('Com borda', bordas)}` +
    `<w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/><w:gridSpan w:val="2"/></w:tcPr>` +
    `${paragraph('Mesclada na horizontal')}</w:tc></w:tr></w:tbl>`

  return zip([
    ['[Content_Types].xml', CONTENT_TYPES.replace(/<Override PartName="\/word\/comments[^>]+>/, '')],
    ['_rels/.rels', ROOT_RELS],
    ['word/document.xml', documentXml(paragraph('Antes da tabela.') + table)],
  ])
}

/**
 * A grid header, like the corpus corporate one: two rows and three columns, the first merged across
 * both.
 */
export async function docxWithHeaderGrid(): Promise<Buffer> {
  const borda =
    `<w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/>` +
    `<w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/>` +
    `<w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders>`

  const celula = (texto: string, largura: string, extra = ''): string =>
    `<w:tc><w:tcPr><w:tcW w:w="${largura}" w:type="dxa"/>${extra}</w:tcPr>` +
    `<w:p><w:r><w:t xml:space="preserve">${texto}</w:t></w:r></w:p></w:tc>`

  const grade =
    `<w:tbl><w:tblPr><w:tblW w:w="10000" w:type="dxa"/>${borda}</w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="8000"/></w:tblGrid>` +
    `<w:tr>${celula('Selo', '2000', '<w:vMerge w:val="restart"/>')}${celula('Chamado 10001', '8000')}</w:tr>` +
    `<w:tr>${celula('', '2000', '<w:vMerge/>')}${celula('Título do documento', '8000')}</w:tr>` +
    // Four rows: the grid exceeds the top margin, and the body must move down.
    `<w:tr>${celula('', '2000', '<w:vMerge/>')}${celula('Terceira linha do cabeçalho', '8000')}</w:tr>` +
    `<w:tr>${celula('', '2000', '<w:vMerge/>')}${celula('Quarta linha do cabeçalho', '8000')}</w:tr>` +
    `</w:tbl>`

  const header = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="${W}">${grade}<w:p/></w:hdr>`

  const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
</Relationships>`

  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
  const corpo =
    paragraph('Primeira linha do corpo.') +
    `<w:sectPr><w:headerReference xmlns:r="${R}" w:type="default" r:id="rId5"/>` +
    `<w:pgSz w:w="11906" w:h="16838"/>` +
    `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708"/>` +
    `</w:sectPr>`

  return zip([
    [
      '[Content_Types].xml',
      CONTENT_TYPES.replace(
        '<Override PartName="/word/document.xml"',
        '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/document.xml"',
      ).replace(/<Override PartName="\/word\/comments[^>]+>/, ''),
    ],
    ['_rels/.rels', ROOT_RELS],
    ['word/_rels/document.xml.rels', rels],
    ['word/header1.xml', header],
    // This document's `w:sectPr` is the fixture's, not the template's empty one.
    [
      'word/document.xml',
      documentXml('').replace('<w:sectPr/>', '').replace('</w:body>', `${corpo}</w:body>`),
    ],
  ])
}

/**
 * Two sheets with a "Página {PAGE}" footer in `w:fldSimple`, as Word writes the number inserted
 * from the band, and `w:pgNumType` in Roman numerals from 3.
 */
export async function docxWithPageNumbering(): Promise<Buffer> {
  const footer = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="${W}"><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t xml:space="preserve">Página </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p></w:ftr>`

  const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
</Relationships>`

  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
  const corpo =
    paragraph('Primeira folha.') +
    '<w:p><w:r><w:br w:type="page"/></w:r></w:p>' +
    paragraph('Segunda folha.') +
    `<w:sectPr><w:footerReference xmlns:r="${R}" w:type="default" r:id="rId5"/>` +
    `<w:pgSz w:w="11906" w:h="16838"/>` +
    `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708"/>` +
    `<w:pgNumType w:fmt="lowerRoman" w:start="3"/>` +
    `</w:sectPr>`

  return zip([
    [
      '[Content_Types].xml',
      CONTENT_TYPES.replace(
        '<Override PartName="/word/document.xml"',
        '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/><Override PartName="/word/document.xml"',
      ).replace(/<Override PartName="\/word\/comments[^>]+>/, ''),
    ],
    ['_rels/.rels', ROOT_RELS],
    ['word/_rels/document.xml.rels', rels],
    ['word/footer1.xml', footer],
    [
      'word/document.xml',
      documentXml('').replace('<w:sectPr/>', '').replace('</w:body>', `${corpo}</w:body>`),
    ],
  ])
}

/**
 * Three sections: portrait with the "Página {PAGE}" footer in Roman numerals; landscape, inheriting
 * the band, ended in a paragraph with text and restarting at 1; and portrait starting on an odd
 * page, where Word inserts a blank sheet.
 */
export async function docxWithSections(): Promise<Buffer> {
  const footer = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="${W}"><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t xml:space="preserve">Página </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p></w:ftr>`

  const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
</Relationships>`

  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
  const margins = `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708"/>`
  const corpo =
    paragraph('Folha em retrato.') +
    `<w:p><w:pPr><w:sectPr><w:footerReference xmlns:r="${R}" w:type="default" r:id="rId5"/>` +
    `<w:pgSz w:w="11906" w:h="16838"/>${margins}<w:pgNumType w:fmt="lowerRoman"/></w:sectPr></w:pPr></w:p>` +
    `<w:p><w:pPr><w:sectPr><w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>${margins}` +
    `<w:pgNumType w:start="1"/></w:sectPr></w:pPr><w:r><w:t xml:space="preserve">Folha em paisagem.</w:t></w:r></w:p>` +
    paragraph('Retrato outra vez, em página ímpar.') +
    `<w:sectPr><w:type w:val="oddPage"/><w:pgSz w:w="11906" w:h="16838"/>${margins}</w:sectPr>`

  return zip([
    [
      '[Content_Types].xml',
      CONTENT_TYPES.replace(
        '<Override PartName="/word/document.xml"',
        '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/><Override PartName="/word/document.xml"',
      ).replace(/<Override PartName="\/word\/comments[^>]+>/, ''),
    ],
    ['_rels/.rels', ROOT_RELS],
    ['word/_rels/document.xml.rels', rels],
    ['word/footer1.xml', footer],
    [
      'word/document.xml',
      documentXml('').replace('<w:sectPr/>', '').replace('</w:body>', `${corpo}</w:body>`),
    ],
  ])
}

/**
 * Columns: a two-column section with a line between them, followed by a one-column continuous
 * section; the two columns balance before it.
 */
export async function docxWithColumns(): Promise<Buffer> {
  const margins = `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708"/>`
  const corpo =
    Array.from({ length: 6 }, (_, index) => paragraph(`Parágrafo ${index + 1} em colunas.`)).join('') +
    `<w:p><w:pPr><w:sectPr><w:pgSz w:w="11906" w:h="16838"/>${margins}` +
    `<w:cols w:num="2" w:space="720" w:sep="1"/></w:sectPr></w:pPr></w:p>` +
    paragraph('Depois das colunas.') +
    `<w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="11906" w:h="16838"/>${margins}</w:sectPr>`

  return zip([
    ['[Content_Types].xml', CONTENT_TYPES.replace(/<Override PartName="\/word\/comments[^>]+>/, '')],
    ['_rels/.rels', ROOT_RELS],
    [
      'word/document.xml',
      documentXml('').replace('<w:sectPr/>', '').replace('</w:body>', `${corpo}</w:body>`),
    ],
  ])
}

/**
 * A screenshot anchored in its own paragraph's place, as LibreOffice writes it: `wp:anchor` without
 * vertical offset, centered, taking height in the flow.
 */
export async function docxWithAnchoredScreenshot(): Promise<Buffer> {
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
  const PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
  const imagem =
    `<w:p><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" ` +
    `relativeHeight="2" behindDoc="0" locked="0" layoutInCell="0" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/>` +
    `<wp:positionH relativeFrom="column"><wp:align>center</wp:align></wp:positionH>` +
    `<wp:positionV relativeFrom="paragraph"><wp:posOffset>635</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="3810000" cy="952500"/><wp:wrapSquare wrapText="bothSides"/>` +
    `<wp:docPr id="1" name="Captura"/>` +
    `<a:graphic><a:graphicData uri="${PIC}"><pic:pic xmlns:pic="${PIC}">` +
    `<pic:nvPicPr><pic:cNvPr id="1" name="Captura"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip xmlns:r="${R}" r:embed="rId9"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3810000" cy="952500"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
    `</pic:pic></a:graphicData></a:graphic>` +
    `</wp:anchor></w:drawing></w:r></w:p>`

  return zip([
    ['[Content_Types].xml', IMAGE_CONTENT_TYPES],
    ['_rels/.rels', ROOT_RELS],
    ['word/_rels/document.xml.rels', IMAGE_RELS],
    ['word/media/quadrado.png', SQUARE_PNG],
    ['word/document.xml', documentXml(paragraph('Antes da captura.') + imagem + paragraph('Depois.'))],
  ])
}

/**
 * The same screenshot in an indented paragraph: in Word it positions itself by the column, and the
 * paragraph indent does not narrow it.
 */
export async function docxWithIndentedScreenshot(): Promise<Buffer> {
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
  const PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture'

  // 3810000 EMU is 100.6 mm: it fits the 146.5 mm column, not the indented paragraph.
  const imagem =
    `<w:p><w:pPr><w:ind w:left="720"/></w:pPr>` +
    `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" ` +
    `relativeHeight="2" behindDoc="0" locked="0" layoutInCell="0" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/>` +
    `<wp:positionH relativeFrom="column"><wp:align>center</wp:align></wp:positionH>` +
    `<wp:positionV relativeFrom="paragraph"><wp:posOffset>635</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="3810000" cy="952500"/><wp:wrapSquare wrapText="bothSides"/>` +
    `<wp:docPr id="1" name="Captura"/>` +
    `<a:graphic><a:graphicData uri="${PIC}"><pic:pic xmlns:pic="${PIC}">` +
    `<pic:nvPicPr><pic:cNvPr id="1" name="Captura"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip xmlns:r="${R}" r:embed="rId9"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3810000" cy="952500"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
    `</pic:pic></a:graphicData></a:graphic>` +
    `</wp:anchor></w:drawing></w:r></w:p>`

  const recuado =
    `<w:p><w:pPr><w:ind w:left="720"/></w:pPr>` +
    `<w:r><w:t xml:space="preserve">Legenda recuada meia polegada.</w:t></w:r></w:p>`

  return zip([
    ['[Content_Types].xml', IMAGE_CONTENT_TYPES],
    ['_rels/.rels', ROOT_RELS],
    ['word/_rels/document.xml.rels', IMAGE_RELS],
    ['word/media/quadrado.png', SQUARE_PNG],
    ['word/document.xml', documentXml(recuado + imagem)],
  ])
}

/**
 * A cover like the manual template's: the title in a positioned box (`wp:anchor` with offset and
 * `wrapNone`), drawn outside the `contenteditable`.
 */
export async function docxWithAnchoredTextBox(): Promise<Buffer> {
  const WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'
  const caixa =
    `<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps">` +
    `<w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" ` +
    `relativeHeight="3" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/>` +
    `<wp:positionH relativeFrom="margin"><wp:posOffset>2160000</wp:posOffset></wp:positionH>` +
    `<wp:positionV relativeFrom="paragraph"><wp:posOffset>360000</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="3800475" cy="2019300"/><wp:wrapNone/>` +
    `<wp:docPr id="7" name="Caixa de Texto"/>` +
    `<a:graphic><a:graphicData uri="${WPS}">` +
    `<wps:wsp><wps:cNvSpPr txBox="1"/>` +
    `<wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3800475" cy="2019300"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>` +
    `<wps:txbx><w:txbxContent><w:p><w:r><w:t>Título da capa</w:t></w:r></w:p></w:txbxContent></wps:txbx>` +
    `<wps:bodyPr rot="0" vert="horz" wrap="square"/></wps:wsp>` +
    `</a:graphicData></a:graphic>` +
    `</wp:anchor></w:drawing></mc:Choice>` +
    // The VML fallback branch, as Word writes it: the same box twice.
    `<mc:Fallback><w:pict xmlns:v="urn:schemas-microsoft-com:vml">` +
    `<v:shape id="Caixa" type="#_x0000_t202" style="position:absolute;width:299pt;height:159pt">` +
    `<v:textbox><w:txbxContent><w:p><w:r><w:t>Título da capa</w:t></w:r></w:p></w:txbxContent></v:textbox>` +
    `</v:shape></w:pict></mc:Fallback>` +
    `</mc:AlternateContent></w:r></w:p>`

  return zip([
    ['[Content_Types].xml', CONTENT_TYPES.replace(/<Override PartName="\/word\/comments[^>]+>/, '')],
    ['_rels/.rels', ROOT_RELS],
    ['word/document.xml', documentXml(caixa + paragraph('Primeiro parágrafo do corpo.'))],
  ])
}

/**
 * A header in a shape group, as in four of the six corpus documents: the title in one box and the
 * page number in another. The box comes whole, because typing in it opens and closes paragraphs.
 */
export async function docxWithHeaderTextBox(): Promise<Buffer> {
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
  const WPG = 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup'

  const forma = (x: string, largura: string, dentro: string): string =>
    `<wps:wsp><wps:cNvSpPr txBox="1"/>` +
    `<wps:spPr><a:xfrm><a:off x="${x}" y="0"/><a:ext cx="${largura}" cy="327600"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>` +
    `<wps:txbx><w:txbxContent>${dentro}</w:txbxContent></wps:txbx>` +
    `<wps:bodyPr/></wps:wsp>`

  // The numbering box carries `PAGE` as `{n}`, and is not editable: it would become a fixed number.
  const numero =
    `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
    `<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>` +
    `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
    `<w:r><w:t>1</w:t></w:r>` +
    `<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`

  const grupo =
    `<w:p><w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" ` +
    `relativeHeight="2" behindDoc="0" locked="0" layoutInCell="0" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/>` +
    `<wp:positionH relativeFrom="page"><wp:posOffset>900000</wp:posOffset></wp:positionH>` +
    `<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="6371640" cy="604440"/><wp:wrapNone/>` +
    `<wp:docPr id="9" name="Grupo do cabeçalho"/>` +
    `<a:graphic><a:graphicData uri="${WPG}"><wpg:wgp xmlns:wpg="${WPG}">` +
    `<wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="6371640" cy="604440"/>` +
    `<a:chOff x="0" y="0"/><a:chExt cx="6371640" cy="604440"/></a:xfrm></wpg:grpSpPr>` +
    forma('0', '900000', numero) +
    forma('1500000', '3052800', `<w:p><w:r><w:t>EVIDÊNCIAS DO ROTEIRO</w:t></w:r></w:p>`) +
    `</wpg:wgp></a:graphicData></a:graphic>` +
    `</wp:anchor></w:drawing></w:r></w:p>`

  const header = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="${W}" xmlns:mc="${MC}" xmlns:wp="${WP}" xmlns:a="${A}" xmlns:wps="${WPS}" mc:Ignorable="wps">${grupo}</w:hdr>`

  const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId5" Type="${R}/header" Target="header1.xml"/>
</Relationships>`

  const corpo =
    paragraph('Primeira linha do corpo.') +
    `<w:sectPr><w:headerReference xmlns:r="${R}" w:type="default" r:id="rId5"/>` +
    `<w:pgSz w:w="11906" w:h="16838"/>` +
    `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708"/>` +
    `</w:sectPr>`

  return zip([
    [
      '[Content_Types].xml',
      CONTENT_TYPES.replace(
        '<Override PartName="/word/document.xml"',
        '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/document.xml"',
      ).replace(/<Override PartName="\/word\/comments[^>]+>/, ''),
    ],
    ['_rels/.rels', ROOT_RELS],
    ['word/_rels/document.xml.rels', rels],
    ['word/header1.xml', header],
    [
      'word/document.xml',
      documentXml('').replace('<w:sectPr/>', '').replace('</w:body>', `${corpo}</w:body>`),
    ],
  ])
}

/**
 * Space on both sides of the joint: Word and LibreOffice add one paragraph's space after to the
 * next one's space before, and CSS keeps the larger margin.
 */
export async function docxWithSpacingOnBothSides(): Promise<Buffer> {
  const espacado = (antes: number, depois: number, texto: string): string =>
    `<w:p><w:pPr><w:spacing w:lineRule="auto" w:line="240" ` +
    `w:before="${antes}" w:after="${depois}"/></w:pPr>` +
    `<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr>` +
    `<w:t xml:space="preserve">${texto}</w:t></w:r></w:p>`

  return zip([
    ['[Content_Types].xml', CONTENT_TYPES.replace(/<Override PartName="\/word\/comments[^>]+>/, '')],
    ['_rels/.rels', ROOT_RELS],
    [
      'word/document.xml',
      documentXml(
        // 283 twips is 14.15 pt on each side of the joint.
        espacado(0, 283, 'Antes da junta.') + espacado(283, 0, 'Depois da junta.'),
      ),
    ],
  ])
}

/**
 * Superscript and subscript (`w:vertAlign`), where formatting changes the meaning: "m2" is not
 * "m²".
 */
export async function docxWithVerticalAlignment(): Promise<Buffer> {
  const run = (align: string, texto: string): string =>
    `<w:r><w:rPr>${align === '' ? '' : `<w:vertAlign w:val="${align}"/>`}</w:rPr>` +
    `<w:t xml:space="preserve">${texto}</w:t></w:r>`

  return zip([
    ['[Content_Types].xml', CONTENT_TYPES.replace(/<Override PartName="\/word\/comments[^>]+>/, '')],
    ['_rels/.rels', ROOT_RELS],
    [
      'word/document.xml',
      documentXml(
        `<w:p>${run('', 'H')}${run('subscript', '2')}${run('', 'O ocupa ')}` +
          `${run('', '18 cm')}${run('superscript', '3')}${run('', '.')}</w:p>`,
      ),
    ],
  ])
}

/**
 * A bullet list declared in `word/numbering.xml`: in the file, sibling paragraphs pointing to a
 * `numId`; in the editor, an element with the items inside.
 */
export async function docxWithBulletList(): Promise<Buffer> {
  const item = (texto: string): string =>
    `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr>` +
    `<w:r><w:t xml:space="preserve">${texto}</w:t></w:r></w:p>`

  const numbering = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="${W}">
<w:abstractNum w:abstractNumId="7"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0A7;"/>
<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>
<w:num w:numId="3"><w:abstractNumId w:val="7"/></w:num>
</w:numbering>`

  return zip([
    [
      '[Content_Types].xml',
      CONTENT_TYPES.replace(
        /<Override PartName="\/word\/comments[^>]+>/,
        '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>',
      ),
    ],
    ['_rels/.rels', ROOT_RELS],
    [
      'word/_rels/document.xml.rels',
      DOCUMENT_RELS.replace(
        /Type="[^"]*comments" Target="comments.xml"/,
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"',
      ),
    ],
    ['word/numbering.xml', numbering],
    [
      'word/document.xml',
      documentXml(paragraph('Antes da lista.') + item('Primeiro item') + item('Segundo item')),
    ],
  ])
}

/**
 * A two-level numbered list, interrupted by a paragraph, and a restart.
 *
 * The second level composes the first in letters (`%1.%2)`), the list continues across the
 * paragraph (same `numId`) and `w:num` 6 is the same definition with `w:startOverride` 10, a
 * separate count. In Word: 1. 1.a) 1.b) 2. | 3. 10.
 */
export async function docxWithMultilevelList(): Promise<Buffer> {
  const item = (texto: string, numId: number, nivel = 0): string =>
    `<w:p><w:pPr><w:numPr><w:ilvl w:val="${nivel}"/><w:numId w:val="${numId}"/></w:numPr></w:pPr>` +
    `<w:r><w:t xml:space="preserve">${texto}</w:t></w:r></w:p>`

  const level = (ilvl: number, fmt: string, text: string, left: number): string =>
    `<w:lvl w:ilvl="${ilvl}"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/><w:lvlText w:val="${text}"/>` +
    `<w:pPr><w:ind w:left="${left}" w:hanging="360"/></w:pPr></w:lvl>`

  const numbering = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="${W}">
<w:abstractNum w:abstractNumId="3">${level(0, 'decimal', '%1.', 360)}${level(1, 'lowerLetter', '%1.%2)', 720)}</w:abstractNum>
<w:num w:numId="5"><w:abstractNumId w:val="3"/></w:num>
<w:num w:numId="6"><w:abstractNumId w:val="3"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="10"/></w:lvlOverride></w:num>
</w:numbering>`

  return zip([
    [
      '[Content_Types].xml',
      CONTENT_TYPES.replace(
        /<Override PartName="\/word\/comments[^>]+>/,
        '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>',
      ),
    ],
    ['_rels/.rels', ROOT_RELS],
    [
      'word/_rels/document.xml.rels',
      DOCUMENT_RELS.replace(
        /Type="[^"]*comments" Target="comments.xml"/,
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"',
      ),
    ],
    ['word/numbering.xml', numbering],
    [
      'word/document.xml',
      documentXml(
        paragraph('Antes da lista.') +
          item('Um', 5) +
          item('Um-a', 5, 1) +
          item('Um-b', 5, 1) +
          item('Dois', 5) +
          paragraph('No meio.') +
          item('Três', 5) +
          item('Dez', 6) +
          paragraph('Depois da lista.'),
      ),
    ],
  ])
}

/**
 * A table with another inside a cell, the only place a block lives in another: width, merge and
 * style do not travel in the model, and what does travel crosses two levels of `content`.
 */
export async function docxWithTable(): Promise<Buffer> {
  const cell = (text: string, width = 4500): string =>
    `<w:tc><w:tcPr><w:tcW w:w="${width}" w:type="dxa"/></w:tcPr>${paragraph(text)}</w:tc>`

  // A `w:tc` ending in `w:tbl` is invalid for Word.
  const inner =
    `<w:tbl><w:tblPr><w:tblStyle w:val="GradeInterna"/></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>` +
    `<w:tr>${cell('Dentro da tabela de dentro', 2000)}</w:tr></w:tbl>`

  const nesting =
    `<w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr>` +
    `${paragraph('Antes da aninhada')}${inner}${paragraph('Depois da aninhada')}</w:tc>`

  const table =
    `<w:tbl><w:tblPr><w:tblStyle w:val="GradeMedia"/><w:tblW w:w="9000" w:type="dxa"/></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid>` +
    `<w:tr><w:trPr><w:tblHeader/></w:trPr>${cell('Cabeçalho A')}${cell('Cabeçalho B')}</w:tr>` +
    `<w:tr>${nesting}${cell('Dado B')}</w:tr></w:tbl>`

  return zip([
    ['[Content_Types].xml', CONTENT_TYPES.replace(/<Override PartName="\/word\/comments[^>]+>/, '')],
    ['_rels/.rels', ROOT_RELS],
    ['word/document.xml', documentXml(paragraph('Antes da tabela.') + table)],
  ])
}

/** A document with nothing the editor does not show. */
export async function docxWithoutExtras(): Promise<Buffer> {
  return zip([
    // The comments content type goes out along with the part.
    ['[Content_Types].xml', CONTENT_TYPES.replace(/<Override PartName="\/word\/comments[^>]+>/, '')],
    ['_rels/.rels', ROOT_RELS],
    ['word/document.xml', documentXml(paragraph('Ata simples.') + paragraph('Sem nada de especial.'))],
  ])
}

const CRC_TABLE = Array.from({ length: 256 }, (_, n) => {
  let c = n
  for (let k = 0; k < 8; k++) c = (c & 1) === 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1
  return c >>> 0
})

function crc32(data: Buffer): number {
  let c = 0xffffffff
  for (const byte of data) c = CRC_TABLE[(c ^ byte) & 0xff]! ^ (c >>> 8)
  return (c ^ 0xffffffff) >>> 0
}

function zip(entries: Array<[string, string | Buffer]>): Buffer {
  const locals: Buffer[] = []
  const central: Buffer[] = []
  let offset = 0

  for (const [name, text] of entries) {
    const filename = Buffer.from(name, 'utf8')
    const data = typeof text === 'string' ? Buffer.from(text, 'utf8') : text
    const checksum = crc32(data)

    const local = Buffer.alloc(30)
    local.writeUInt32LE(0x04034b50, 0)
    local.writeUInt16LE(20, 4) // minimum version
    // Method 0 = stored. No date, so the fixture is reproducible.
    local.writeUInt32LE(checksum, 14)
    local.writeUInt32LE(data.length, 18)
    local.writeUInt32LE(data.length, 22)
    local.writeUInt16LE(filename.length, 26)

    const entry = Buffer.concat([local, filename, data])
    locals.push(entry)

    const header = Buffer.alloc(46)
    header.writeUInt32LE(0x02014b50, 0)
    header.writeUInt16LE(20, 4)
    header.writeUInt16LE(20, 6)
    header.writeUInt32LE(checksum, 16)
    header.writeUInt32LE(data.length, 20)
    header.writeUInt32LE(data.length, 24)
    header.writeUInt16LE(filename.length, 28)
    header.writeUInt32LE(offset, 42)
    central.push(Buffer.concat([header, filename]))

    offset += entry.length
  }

  const directory = Buffer.concat(central)
  const end = Buffer.alloc(22)
  end.writeUInt32LE(0x06054b50, 0)
  end.writeUInt16LE(entries.length, 8)
  end.writeUInt16LE(entries.length, 10)
  end.writeUInt32LE(directory.length, 12)
  end.writeUInt32LE(offset, 16)

  return Buffer.concat([...locals, directory, end])
}

/**
 * By the central directory, not the local headers: LibreOffice writes the sizes after the data, and
 * the local header carries zero.
 */
export async function entryOf(path: string, name: string): Promise<string> {
  const { readFile } = await import('node:fs/promises')
  const { inflateRawSync } = await import('node:zlib')
  const zip = await readFile(path)

  const directoryEnd = zip.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]))
  const entries = zip.readUInt16LE(directoryEnd + 10)
  let position = zip.readUInt32LE(directoryEnd + 16)
  for (let index = 0; index < entries; index++) {
    const method = zip.readUInt16LE(position + 10)
    const compressedSize = zip.readUInt32LE(position + 20)
    const nameLength = zip.readUInt16LE(position + 28)
    const extraLength = zip.readUInt16LE(position + 30)
    const commentLength = zip.readUInt16LE(position + 32)
    const localHeader = zip.readUInt32LE(position + 42)
    if (zip.toString('utf8', position + 46, position + 46 + nameLength) === name) {
      const start = localHeader + 30 + zip.readUInt16LE(localHeader + 26) + zip.readUInt16LE(localHeader + 28)
      const data = zip.subarray(start, start + compressedSize)
      return (method === 0 ? data : inflateRawSync(data)).toString('utf8')
    }
    position += 46 + nameLength + extraLength + commentLength
  }

  throw new Error(`${name} não encontrado em ${path}`)
}

/**
 * Styles as in the corpus: the translated id (`Ttulo1`, as Word writes it) and the internal name
 * not (`heading 1`), and a paragraph in the author's style (`Citao`), which the pane shows as the
 * cursor's style.
 */
export async function docxWithNamedStyles(extraParagraphs = ''): Promise<Buffer> {
  const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${W}">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
<w:style w:type="paragraph" w:styleId="Ttulo1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/>
<w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="Citao"><w:name w:val="Citação recuada"/><w:basedOn w:val="Normal"/><w:qFormat/>
<w:pPr><w:ind w:left="720"/></w:pPr><w:rPr><w:i/></w:rPr></w:style>
<w:style w:type="character" w:styleId="FonteParagrPadro" w:default="1"><w:name w:val="Default Paragraph Font"/><w:semiHidden/><w:unhideWhenUsed/></w:style>
</w:styles>`

  const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`

  const comEstilo = (id: string, texto: string): string =>
    `<w:p><w:pPr><w:pStyle w:val="${id}"/></w:pPr><w:r><w:t xml:space="preserve">${texto}</w:t></w:r></w:p>`

  return zip([
    [
      '[Content_Types].xml',
      CONTENT_TYPES.replace(
        '<Override PartName="/word/document.xml"',
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/document.xml"',
      ).replace(/<Override PartName="\/word\/comments[^>]+>/, ''),
    ],
    ['_rels/.rels', ROOT_RELS],
    ['word/_rels/document.xml.rels', rels],
    ['word/styles.xml', styles],
    [
      'word/document.xml',
      documentXml(
        comEstilo('Ttulo1', 'Relatório anual') + comEstilo('Citao', 'Um trecho citado.') + extraParagraphs,
      ),
    ],
  ])
}

/**
 * The same styles with direct formatting on top: zero indent on the quote, "keep with next" turned
 * off, partial spacing and line spacing, and the paragraph mark font. That is all the block
 * carries.
 */
export async function docxWithDirectOverStyles(): Promise<Buffer> {
  const p = (props: string, text: string): string =>
    `<w:p><w:pPr>${props}</w:pPr><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`
  return docxWithNamedStyles(
    p('<w:pStyle w:val="Citao"/><w:keepNext w:val="0"/><w:ind w:left="0"/>', 'Citação sem recuo.') +
      p('<w:spacing w:after="0"/><w:jc w:val="center"/>', 'Só o espaço depois.') +
      p('<w:spacing w:line="360" w:lineRule="auto"/>', 'Uma linha e meia.') +
      p('<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr>', 'Marca em Arial.') +
      p('<w:pStyle w:val="EstiloQueNaoExiste"/>', 'Estilo que o documento não define.'),
  )
}

/**
 * Word's references as it writes them, the same document as `Fixtures.WithReferences`: a table of
 * contents in a content control with `TOC` across paragraphs and `PAGEREF` in the links, `_Toc…`, a
 * bookmark with its end in the body, `SEQ`, `REF` and an internal link.
 */
export const REFERENCES_BODY =
  '<w:sdt><w:sdtPr><w:id w:val="-1"/><w:docPartObj><w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtEndPr/><w:sdtContent><w:p><w:pPr><w:pStyle w:val="CabealhodoSumrio"/></w:pPr><w:r><w:t>Sumário</w:t></w:r></w:p>' +
  '<w:p><w:pPr><w:pStyle w:val="Sumrio1"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9016"/></w:tabs></w:pPr><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:hyperlink w:anchor="_Toc100" w:history="1"><w:r><w:t>Introdução</w:t></w:r><w:r><w:tab/></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGEREF _Toc100 \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:hyperlink></w:p>' +
  '<w:p><w:pPr><w:pStyle w:val="Sumrio2"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9016"/></w:tabs></w:pPr><w:hyperlink w:anchor="_Toc101" w:history="1"><w:r><w:t>Escopo</w:t></w:r><w:r><w:tab/></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGEREF _Toc101 \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:hyperlink></w:p>' +
  '<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>' +
  '</w:sdtContent></w:sdt><w:p><w:pPr><w:pStyle w:val="Ttulo1"/></w:pPr><w:bookmarkStart w:id="0" w:name="_Toc100"/><w:r><w:t>Introdução</w:t></w:r><w:bookmarkEnd w:id="0"/></w:p>' +
  '<w:p><w:bookmarkStart w:id="1" w:name="Resumo"/><w:r><w:t xml:space="preserve">O resumo começa aqui </w:t></w:r></w:p>' +
  '<w:p><w:r><w:t>e termina aqui.</w:t></w:r></w:p>' +
  '<w:bookmarkEnd w:id="1"/><w:p><w:pPr><w:pStyle w:val="Ttulo2"/></w:pPr><w:bookmarkStart w:id="2" w:name="_Toc101"/><w:r><w:t>Escopo</w:t></w:r><w:bookmarkEnd w:id="2"/></w:p>' +
  '<w:p><w:pPr><w:pStyle w:val="Legenda"/></w:pPr><w:bookmarkStart w:id="3" w:name="_Ref200"/><w:r><w:t xml:space="preserve">Figura </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> SEQ Figura \\* ARABIC </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:noProof/></w:rPr><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:bookmarkEnd w:id="3"/><w:r><w:t xml:space="preserve"> — Arquitetura</w:t></w:r></w:p>' +
  '<w:p><w:r><w:t xml:space="preserve">Como mostra a </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> REF _Ref200 \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Figura 1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t xml:space="preserve">, na página </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGEREF _Ref200 \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:t xml:space="preserve">. Veja o </w:t></w:r><w:hyperlink w:anchor="Resumo" w:history="1"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>resumo</w:t></w:r></w:hyperlink><w:r><w:t>.</w:t></w:r></w:p>'

const REFERENCES_STYLES =
  '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Ttulo1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Ttulo2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:sz w:val="26"/></w:rPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="CabealhodoSumrio"><w:name w:val="TOC Heading"/><w:basedOn w:val="Ttulo1"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:outlineLvl w:val="9"/></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Sumrio1"><w:name w:val="toc 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:spacing w:after="100"/></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Sumrio2"><w:name w:val="toc 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:pPr><w:spacing w:after="100"/><w:ind w:left="220"/></w:pPr></w:style>' +
  '<w:style w:type="paragraph" w:styleId="Legenda"><w:name w:val="caption"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:spacing w:after="200"/></w:pPr><w:rPr><w:i/><w:sz w:val="18"/></w:rPr></w:style>' +
  '<w:style w:type="character" w:default="1" w:styleId="Fontepargpadro"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/><w:unhideWhenUsed/></w:style>' +
  '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:basedOn w:val="Fontepargpadro"/><w:uiPriority w:val="99"/><w:unhideWhenUsed/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>'

export async function docxWithReferences(body = REFERENCES_BODY): Promise<Buffer> {
  const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:styles xmlns:w="${W}">${REFERENCES_STYLES}</w:styles>`
  const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`

  return zip([
    [
      '[Content_Types].xml',
      CONTENT_TYPES.replace(
        '<Override PartName="/word/document.xml"',
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/document.xml"',
      ).replace(/<Override PartName="\/word\/comments[^>]+>/, ''),
    ],
    ['_rels/.rels', ROOT_RELS],
    ['word/_rels/document.xml.rels', rels],
    ['word/styles.xml', styles],
    ['word/document.xml', documentXml(body)],
  ])
}

/** A long table on A4 with 25 mm margins: a single structure, many sheets. */
export async function docxWithLongTable(rows = 80, header = false): Promise<Buffer> {
  // With `header`, the first row is the one Word repeats at the top of each sheet.
  const table =
    '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>' +
    (header
      ? `<w:tr><w:trPr><w:tblHeader/></w:trPr><w:tc>${paragraph('Cabeçalho repetido')}</w:tc></w:tr>`
      : '') +
    Array.from(
      { length: rows },
      (_, index) => `<w:tr><w:tc>${paragraph(`Linha ${index + 1} da tabela longa`)}</w:tc></w:tr>`,
    ).join('') +
    '</w:tbl>'
  return zip([
    ['[Content_Types].xml', CONTENT_TYPES.replace(/<Override PartName="\/word\/comments[^>]+>/, '')],
    ['_rels/.rels', ROOT_RELS],
    [
      'word/document.xml',
      documentXml(table).replace(
        '<w:sectPr/>',
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417"/></w:sectPr>',
      ),
    ],
  ])
}

/**
 * A long document with one footnote per paragraph: notes take height from the sheets and push lines
 * to the next. A4 paper, margins and font declared, so the sheet count also holds in LibreOffice.
 */
export async function docxWithManyFootnotes(count = 24): Promise<Buffer> {
  const text =
    'Os conselheiros discutiram o orçamento do ano seguinte, as obras da sede e a contratação de pessoal, ' +
    'e decidiram adiar a votação até que os números da tesouraria fossem revistos por uma comissão própria.'
  const noteText =
    'Conforme a ata anterior, a comissão terá trinta dias para revisar os números e apresentar um parecer ' +
    'escrito, que será lido na abertura da reunião seguinte, antes de qualquer outro assunto da pauta.'
  let body = ''
  let notes = ''
  for (let id = 1; id <= count; id++) {
    body +=
      `<w:p><w:r><w:t xml:space="preserve">${id}. ${text}</w:t></w:r>` +
      `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="${id}"/></w:r>` +
      `<w:r><w:t xml:space="preserve"> ${text}</w:t></w:r></w:p>`
    notes +=
      `<w:footnote w:id="${id}"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>` +
      `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>` +
      `<w:r><w:t xml:space="preserve"> Nota ${id}. ${noteText}</w:t></w:r></w:p></w:footnote>`
  }
  return footnotesPackage(body, notes)
}

/** The package with the given body and footnotes, in the note tests' font and paper. */
function footnotesPackage(body: string, notes: string): Buffer {
  const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${W}">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif"/><w:sz w:val="24"/></w:rPr></w:rPrDefault>
<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>
<w:style w:type="paragraph" w:styleId="FootnoteText"><w:name w:val="footnote text"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="20"/></w:rPr></w:style>
<w:style w:type="character" w:styleId="FootnoteReference"><w:name w:val="footnote reference"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>
</w:styles>`
  const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
</Relationships>`
  const footnotes =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:footnotes xmlns:w="${W}">` +
    `<w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote>` +
    `<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>` +
    notes +
    `</w:footnotes>`
  const document = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W}"><w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1417" w:right="1701" w:bottom="1417" w:left="1701" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr></w:body></w:document>`

  return zip([
    [
      '[Content_Types].xml',
      CONTENT_TYPES.replace(
        /<Override PartName="\/word\/comments[^>]+>/,
        '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
          '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>',
      ),
    ],
    ['_rels/.rels', ROOT_RELS],
    ['word/_rels/document.xml.rels', rels],
    ['word/styles.xml', styles],
    ['word/document.xml', document],
    ['word/footnotes.xml', footnotes],
  ])
}

/**
 * A footnote taller than a sheet (`lines` short paragraphs) on the first line, and a few paragraphs
 * after: the continuation goes on to the following sheets.
 */
export async function docxWithLongFootnote(lines = 70): Promise<Buffer> {
  let note =
    `<w:footnote w:id="1"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>` +
    `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>` +
    `<w:r><w:t xml:space="preserve"> Nota longa.</w:t></w:r></w:p>`
  for (let line = 0; line < lines; line++) {
    note += `<w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:t>Linha ${line} da nota longa.</w:t></w:r></w:p>`
  }
  note += '</w:footnote>'
  const body =
    `<w:p><w:r><w:t>Alfa</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteReference w:id="1"/></w:r></w:p>` +
    ['Beta', 'Gama', 'Delta'].map((word) => `<w:p><w:r><w:t>${word}</w:t></w:r></w:p>`).join('')
  return footnotesPackage(body, note)
}
