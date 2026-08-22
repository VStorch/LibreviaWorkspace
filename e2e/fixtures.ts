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

function zip(entries: Array<[string, string]>): Buffer {
  const locals: Buffer[] = []
  const central: Buffer[] = []
  let offset = 0

  for (const [name, text] of entries) {
    const filename = Buffer.from(name, 'utf8')
    const data = Buffer.from(text, 'utf8')
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
