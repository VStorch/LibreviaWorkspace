/**
 * Main's path when opening and saving `.docx`, against the corpus in `LIBREVIA_CORPUS_DIR`, which
 * is not in the repository; skipped without the variable. CI covers the same structures with the
 * sidecar fixtures.
 */

import { access, constants, mkdtemp, readdir, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { inflateRawSync } from 'node:zlib'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import { SidecarClient } from '../sidecar/client.js'
import { sidecarPathIn } from '../sidecar/locate.js'
import { adoptDocxOriginal, followDocxOriginal, forgetOpenedDocx, openDocx, saveDocx } from './index.js'

const corpusDirectory = process.env['LIBREVIA_CORPUS_DIR']

const documents =
  corpusDirectory === undefined
    ? []
    : (await readdir(corpusDirectory))
        .filter((name) => name.toLowerCase().endsWith('.docx'))
        .map((name) => join(corpusDirectory, name))

const client = new SidecarClient(() => Promise.resolve(sidecarPathIn(process.cwd())))
afterAll(() => client.dispose())

const published = await access(sidecarPathIn(process.cwd()), constants.X_OK).then(
  () => true,
  () => false,
)

const page = {
  size: 'A4',
  orientation: 'portrait',
  margins: { top: 25, right: 25, bottom: 25, left: 25 },
  header: '',
  footer: '',
  headerBand: null,
  footerBand: null,
}

function newDocument(text: string): string {
  return JSON.stringify({
    page,
    doc: { type: 'doc', content: [{ type: 'paragraph', content: [{ type: 'text', text }] }] },
  })
}

/** A one-pixel PNG, the smallest the image writer accepts. */
const onePixelPng =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFAAH/q842iQAAAABJRU5ErkJggg=='

/** A list and an image, the two parts that accumulate. */
function newDocumentWithListAndImage(text: string): string {
  return JSON.stringify({
    page,
    doc: {
      type: 'doc',
      content: [
        { type: 'paragraph', content: [{ type: 'text', text }] },
        {
          type: 'bulletList',
          content: [
            { type: 'listItem', content: [{ type: 'paragraph', content: [{ type: 'text', text: 'um' }] }] },
            { type: 'listItem', content: [{ type: 'paragraph', content: [{ type: 'text', text: 'dois' }] }] },
          ],
        },
        { type: 'image', attrs: { src: onePixelPng } },
      ],
    },
  })
}

describe.skipIf(documents.length === 0)('corpus real', () => {
  it.each(documents)('abre %s', async (path) => {
    const opened = await openDocx(client, path)
    const model = JSON.parse(opened.content) as {
      format: string
      page: { size: string; margins: Record<string, number> }
      doc: { type: string; content?: unknown[] }
    }

    expect(model.format).toBe('sdoc')
    expect(model.doc.type).toBe('doc')
    expect(model.doc.content?.length).toBeGreaterThan(0)
    expect(['A4', 'Letter']).toContain(model.page.size)
  })

  it.each(documents)('salva %s sem reescrever nada quando nada foi editado', async (path) => {
    // Opening and saving by reflex costs nothing.
    forgetOpenedDocx()
    const opened = await openDocx(client, path)
    const saved = await saveDocx(client, opened.content, { origin: path, destination: path })

    expect(saved.bytes.length).toBeGreaterThan(0)
    expect(saved.inventory.lost).toEqual([])
  })

  it('preserva todas as partes do pacote menos o corpo', async () => {
    const path = documents[0]!
    forgetOpenedDocx()

    const original = await readFile(path)
    const opened = await openDocx(client, path)
    const saved = await saveDocx(client, opened.content, { origin: path, destination: path })

    const before = await listParts(original)
    const after = await listParts(Buffer.from(saved.bytes))

    expect([...after.keys()].sort()).toEqual([...before.keys()].sort())

    const changed = [...before.keys()].filter((name) => !before.get(name)!.equals(after.get(name)!))
    expect(changed).toEqual(['word/document.xml'])
  })

  it('não empresta o pacote aberto a um documento de outra origem', async () => {
    // A new document after opening a `.docx` does not write over its bytes.
    const path = documents[0]!
    forgetOpenedDocx()
    await openDocx(client, path)

    const saved = await saveDocx(client, newDocument('Outro documento.'), {
      origin: null,
      destination: '/tmp/novo.docx',
    })
    const parts = await listParts(Buffer.from(saved.bytes))

    expect([...parts.keys()].some((name) => name.startsWith('word/header'))).toBe(false)
  })
})

/**
 * A document born in the editor saved as `.docx`, over the package the sidecar creates. Depends on
 * the published sidecar, not on the corpus.
 */
describe.skipIf(!published)('documento novo em .docx', () => {
  it('grava sobre o pacote mínimo e o texto volta ao reabrir', async () => {
    const saved = await saveDocx(client, newDocument('Texto do documento novo.'), {
      origin: null,
      destination: '/tmp/novo.docx',
    })

    const parts = await listParts(Buffer.from(saved.bytes))
    expect(parts.has('word/styles.xml')).toBe(true)
    expect(parts.has('word/numbering.xml')).toBe(false)
    expect(saved.inventory.lost).toEqual([])

    // The original kept is the minimal package, not the saved bytes.
    const keptParts = await listParts(saved.original)
    expect(keptParts.get('word/document.xml')!.toString('utf8')).not.toContain('Texto do documento novo.')
  })

  it('a segunda gravação no mesmo arquivo parte do mesmo pacote que a primeira', async () => {
    const destination = '/tmp/novo-duas-vezes.docx'
    forgetOpenedDocx()
    const first = await saveDocx(client, newDocument('Mesma versão.'), { origin: null, destination })
    followDocxOriginal(null, destination, first)

    const second = await saveDocx(client, newDocument('Mesma versão.'), {
      origin: destination,
      destination,
    })

    // Same package and same content, same bytes.
    expect(Buffer.compare(second.original, first.original)).toBe(0)
    expect(Buffer.compare(Buffer.from(second.bytes), Buffer.from(first.bytes))).toBe(0)
  })

  it('gravar cinco vezes o mesmo documento novo não acumula partes', async () => {
    // Always starting from the minimal package, saving again adds no numbering or image to what was
    // there.
    const destination = '/tmp/novo-cinco-vezes.docx'
    forgetOpenedDocx()

    const counted: { media: number; numbering: number }[] = []
    for (let round = 0; round < 5; round++) {
      const origin = round === 0 ? null : destination
      const saved = await saveDocx(client, newDocumentWithListAndImage(`Versão ${round + 1}.`), {
        origin,
        destination,
      })
      followDocxOriginal(origin, destination, saved)

      const parts = await listParts(Buffer.from(saved.bytes))
      const numbering = parts.get('word/numbering.xml')?.toString('utf8') ?? ''
      counted.push({
        media: [...parts.keys()].filter((name) => name.includes('media/')).length,
        numbering: numbering.split('<w:abstractNum ').length - 1,
      })
    }

    // And it is not empty: the list and the image are in the package.
    expect(counted[0]).toEqual({ media: 1, numbering: 1 })
    expect(counted).toEqual(counted.map(() => counted[0]))
  })

  it('declara a perda do pacote de origem quando o modelo tem oid e o original não está aqui', async () => {
    // A model with `oid` saved without its package loses styles, notes and comments: the user has
    // to read that on screen.
    forgetOpenedDocx()
    const content = JSON.stringify({
      page,
      doc: {
        type: 'doc',
        content: [
          { type: 'paragraph', attrs: { oid: 'b1' }, content: [{ type: 'text', text: 'Veio de um .docx.' }] },
        ],
      },
    })

    const saved = await saveDocx(client, content, {
      origin: '/tmp/origem.docx',
      destination: '/tmp/destino.docx',
    })

    expect(saved.inventory.lost).toContain(
      'estilos, numeração das notas e demais partes do arquivo .docx de origem',
    )
  })

  it('avisa quando as faixas de um .docx de origem não têm onde ser gravadas', async () => {
    // A `.sdoc` that was a `.docx` carries the source file's headers and footers.
    const content = JSON.stringify({
      page: { ...page, headerBand: { left: [], center: [], right: [], rule: true, rows: [] } },
      doc: { type: 'doc', content: [] },
    })

    const saved = await saveDocx(client, content, { origin: null, destination: '/tmp/faixa.docx' })

    expect(saved.inventory.lost).toContain('cabeçalho e rodapé do arquivo .docx de origem')
  })
})

/**
 * The path that recognizes the original package comes from the dialog on open and from the
 * renderer's `origin` on save: they must match, or the opened `.docx` is written over the minimal
 * package.
 */
describe.skipIf(!published)('o caminho do pacote original', () => {
  let directory = ''

  beforeAll(async () => {
    directory = await mkdtemp(join(tmpdir(), 'librevia-caminho-'))
  })
  afterAll(() => rm(directory, { recursive: true, force: true }))

  /** Written by the writer itself. */
  async function docxOnDisk(name: string): Promise<string> {
    const path = join(directory, name)
    forgetOpenedDocx()
    const created = await saveDocx(client, newDocument('Documento de origem.'), {
      origin: null,
      destination: path,
    })
    await writeFile(path, created.bytes)
    return path
  }

  /** The same file spelled differently, without `join`, which would normalize it. */
  function detoured(path: string): string {
    return `${directory}/.//${path.slice(directory.length + 1)}`
  }

  it('reconhece o original quando a abertura veio por um caminho escrito de outro jeito', async () => {
    const path = await docxOnDisk('aberto.docx')
    forgetOpenedDocx()

    const opened = await openDocx(client, detoured(path))
    const saved = await saveDocx(client, opened.content, { origin: path, destination: path })

    expect(saved.inventory.lost).toEqual([])
    expect(Buffer.compare(saved.original, await readFile(path))).toBe(0)
  })

  it('reconhece o original quando o origin da gravação vem escrito de outro jeito', async () => {
    const path = await docxOnDisk('gravado.docx')
    forgetOpenedDocx()

    const opened = await openDocx(client, path)
    const saved = await saveDocx(client, opened.content, {
      origin: detoured(path),
      destination: path,
    })

    expect(saved.inventory.lost).toEqual([])
    expect(Buffer.compare(saved.original, await readFile(path))).toBe(0)
  })

  it('reata o original recuperado por um caminho escrito de outro jeito', async () => {
    // Recovery rereads the bytes from disk through the draft's path.
    const path = await docxOnDisk('recuperado.docx')
    forgetOpenedDocx()

    // The model comes from the draft; here, from a read that is forgotten right away.
    const opened = await openDocx(client, path)
    forgetOpenedDocx()

    expect(await adoptDocxOriginal(detoured(path))).toBe(true)
    const saved = await saveDocx(client, opened.content, { origin: path, destination: path })

    expect(saved.inventory.lost).toEqual([])
    expect(Buffer.compare(saved.original, await readFile(path))).toBe(0)
  })
})

/** Through the central directory, including entries with a data descriptor. */
async function listParts(zip: Buffer): Promise<Map<string, Buffer>> {
  let end = -1
  for (let at = zip.length - 22; at >= Math.max(0, zip.length - 65_557); at--) {
    if (zip.readUInt32LE(at) === 0x06054b50) {
      end = at
      break
    }
  }
  if (end < 0) throw new Error('ZIP end of central directory not found')

  const parts = new Map<string, Buffer>()
  let at = zip.readUInt32LE(end + 16)
  const count = zip.readUInt16LE(end + 10)
  for (let entry = 0; entry < count; entry++) {
    if (zip.readUInt32LE(at) !== 0x02014b50) throw new Error('Invalid ZIP central directory')
    const method = zip.readUInt16LE(at + 10)
    const compressed = zip.readUInt32LE(at + 20)
    const nameLength = zip.readUInt16LE(at + 28)
    const extraLength = zip.readUInt16LE(at + 30)
    const commentLength = zip.readUInt16LE(at + 32)
    const local = zip.readUInt32LE(at + 42)
    const name = zip.subarray(at + 46, at + 46 + nameLength).toString('utf8')
    if (!name.endsWith('/')) {
      if (zip.readUInt32LE(local) !== 0x04034b50) throw new Error('Invalid ZIP local entry')
      const start = local + 30 + zip.readUInt16LE(local + 26) + zip.readUInt16LE(local + 28)
      const data = zip.subarray(start, start + compressed)
      if (method !== 0 && method !== 8) throw new Error(`Unsupported ZIP method ${method}`)
      parts.set(name, method === 0 ? data : inflateRawSync(data))
    }
    at += 46 + nameLength + extraLength + commentLength
  }
  return parts
}
