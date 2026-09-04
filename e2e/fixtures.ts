/**
 * Documentos `.docx` construídos em código, não versionados como binário.
 *
 * Mesmo princípio dos fixtures do sidecar: o que o documento contém fica
 * legível na revisão. Um `.docx` no git é uma caixa preta que ninguém confere —
 * e aqui o conteúdo é justamente o que está sendo testado.
 *
 * O ZIP é escrito à mão, com os arquivos **armazenados** e não comprimidos.
 * Sem compressão o formato é só cabeçalho, dados e um índice no fim: cabe em
 * meia página e não traz dependência nenhuma. O tamanho não importa — estes
 * documentos vivem alguns segundos dentro de uma pasta temporária.
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

/** O parágrafo que ancora o comentário — é ele que torna a perda possível. */
const COMMENTED_PARAGRAPH =
  `<w:p><w:commentRangeStart w:id="1"/>` +
  `<w:r><w:t xml:space="preserve">Segundo parágrafo, com um comentário ancorado.</w:t></w:r>` +
  `<w:commentRangeEnd w:id="1"/><w:r><w:commentReference w:id="1"/></w:r></w:p>`

function documentXml(body: string): string {
  // Os namespaces de desenho entram sempre: declarar só quando o corpo os usa
  // faria cada fixture novo redescobrir por que o pacote não abre.
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W}" xmlns:mc="${MC}" xmlns:wp="${WP}" xmlns:a="${A}" xmlns:wps="${WPS}" mc:Ignorable="wps"><w:body>${body}<w:sectPr/></w:body></w:document>`
}

const COMMENTS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="${W}">
<w:comment w:id="1" w:author="Revisor" w:date="2026-01-01T00:00:00Z"><w:p><w:r><w:t>Conferir este número.</w:t></w:r></w:p></w:comment>
</w:comments>`

/** Documento com um comentário ancorado no segundo parágrafo. */
export async function docxWithComment(): Promise<Buffer> {
  return zip([
    ['[Content_Types].xml', CONTENT_TYPES],
    ['_rels/.rels', ROOT_RELS],
    ['word/_rels/document.xml.rels', DOCUMENT_RELS],
    [
      'word/document.xml',
      documentXml(paragraph('Ata da reunião de terça.') + COMMENTED_PARAGRAPH + paragraph('Fim.')),
    ],
    ['word/comments.xml', COMMENTS_XML],
  ])
}

/**
 * Documento com uma caixa de texto ancorada.
 *
 * É o que a regeneração destrói: o editor mostra o texto de dentro, mas não
 * sabe redesenhar a forma, então um bloco reescrito volta para o arquivo sem
 * ela. Serve de sentinela para a gravação cirúrgica — se a caixa sobreviveu, o
 * XML original do bloco foi preservado.
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

/**
 * Um PNG quadrado de 4 × 4, do tamanho de um comentário.
 *
 * Quadrado de propósito: o documento que o usa pede uma caixa de 400 × 100, e é
 * a divergência entre a proporção do arquivo e a que o documento pede que o
 * teste observa.
 */
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
 * Documento com uma imagem esticada no fluxo do texto.
 *
 * O `.docx` diz em `wp:extent` de que tamanho a imagem é **na página**, e esse
 * tamanho não precisa ter a proporção do arquivo: quem arrasta um canto sem
 * travar a proporção estica a imagem, e o Word desenha esticado. Aqui um PNG
 * quadrado é declarado como 400 × 100 px — 3810000 × 952500 EMU.
 */
export async function docxWithStretchedImage(): Promise<Buffer> {
  const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
  const PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
  const imagem =
    `<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">` +
    `<wp:extent cx="3810000" cy="952500"/><wp:docPr id="1" name="Quadrado"/>` +
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

/**
 * Documento cujo cabeçalho é uma grade, como o cabeçalho corporativo do corpus.
 *
 * Duas linhas e três colunas, com a primeira coluna mesclada pelas duas linhas.
 * É a estrutura que, achatada em esquerda-centro-direita, virava uma fileira de
 * palavras por cima da primeira linha do texto.
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
    // Quatro linhas de propósito: assim a grade fica mais alta que a margem de
    // cima, que é a situação em que o corpo tem de descer para debaixo dela.
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
    // O `w:sectPr` deste documento é o do fixture, não o vazio do molde.
    [
      'word/document.xml',
      documentXml('').replace('<w:sectPr/>', '').replace('</w:body>', `${corpo}</w:body>`),
    ],
  ])
}

/**
 * Documento com uma imagem **ancorada** no lugar do próprio parágrafo.
 *
 * É como o LibreOffice grava captura de tela: `wp:anchor` sem deslocamento
 * vertical, centralizada na coluna. Tratá-la como posição na folha fazia a
 * imagem deixar de ocupar altura — o texto se fechava por cima dela, e um
 * documento de trinta capturas encolhia de quinze folhas para quatro.
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
 * Capa como a do modelo de manual: o título mora numa caixa **posicionada**.
 *
 * `wp:anchor` com deslocamento de verdade e `wrapNone` — o texto não está no
 * fluxo, está numa posição da folha. É o caso em que a pessoa olha a capa,
 * quer trocar o título e não tem onde clicar, porque a caixa é desenhada fora
 * do `contenteditable`.
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
    // O ramo de reserva é o VML antigo, como o Word grava: a mesma caixa
    // escrita duas vezes. Salvar só uma deixaria o arquivo dizendo duas
    // coisas, e qual delas aparece dependeria de quem abre.
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
 * Dois parágrafos que pedem espaço nos dois lados da junta.
 *
 * O Word e o LibreOffice **somam** o espaço depois de um com o espaço antes do
 * seguinte; o CSS funde as duas margens e fica com a maior. Meia linha por
 * junta, e ela se acumula até a folha cortar noutro lugar.
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
        // 283 twips são 14,15 pt de cada lado da junta.
        espacado(0, 283, 'Antes da junta.') + espacado(283, 0, 'Depois da junta.'),
      ),
    ],
  ])
}

/** Documento sem nada que o editor não mostre. */
export async function docxWithoutExtras(): Promise<Buffer> {
  return zip([
    // O tipo de conteúdo dos comentários fica de fora junto com a parte.
    ['[Content_Types].xml', CONTENT_TYPES.replace(/<Override PartName="\/word\/comments[^>]+>/, '')],
    ['_rels/.rels', ROOT_RELS],
    ['word/document.xml', documentXml(paragraph('Ata simples.') + paragraph('Sem nada de especial.'))],
  ])
}

// --- ZIP mínimo, com entradas armazenadas ----------------------------------

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
    local.writeUInt16LE(20, 4) // versão mínima
    // Método 0 = armazenado. Sem data: o conteúdo é o que importa, e uma data
    // fixa deixa o fixture reprodutível byte a byte.
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
 * O conteúdo de uma parte de dentro de um `.docx`, sem descompactar em disco.
 *
 * O `.docx` que o aplicativo grava tem as entradas comprimidas; o que estes
 * fixtures produzem, não. O leitor cobre os dois porque é o mesmo teste que
 * escreve um e lê o outro.
 */
export async function entryOf(path: string, name: string): Promise<string> {
  const { readFile } = await import('node:fs/promises')
  const { promisify } = await import('node:util')
  const { inflateRaw } = await import('node:zlib')
  const inflate = promisify(inflateRaw)
  const zip = await readFile(path)

  for (let i = 0; i + 30 <= zip.length; i++) {
    if (zip.readUInt32LE(i) !== 0x04034b50) continue

    const metodo = zip.readUInt16LE(i + 8)
    const comprimido = zip.readUInt32LE(i + 18)
    const original = zip.readUInt32LE(i + 22)
    const tamanhoNome = zip.readUInt16LE(i + 26)
    const extra = zip.readUInt16LE(i + 28)
    if (zip.subarray(i + 30, i + 30 + tamanhoNome).toString('utf8') !== name) continue

    const fim = i + 30 + tamanhoNome + extra + (comprimido > 0 ? comprimido : original)
    const dados = zip.subarray(i + 30 + tamanhoNome + extra, fim)
    return (metodo === 0 ? dados : await inflate(dados)).toString('utf8')
  }

  throw new Error(`${name} não encontrado em ${path}`)
}
