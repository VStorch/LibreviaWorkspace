/**
 * O caminho do processo main ao abrir e salvar `.docx`, contra o corpus de
 * `LIBREVIA_CORPUS_DIR`, que não entra no repositório; pulado sem a variável. O CI
 * cobre as mesmas estruturas com os fixtures do sidecar.
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

/** Um PNG de um pixel, o menor que o escritor de imagem aceita. */
const onePixelPng =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFAAH/q842iQAAAABJRU5ErkJggg=='

/** Documento novo com uma lista e uma imagem — as duas partes que se acumulam. */
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
    // Abrir e salvar por reflexo custa zero.
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
    // O documento novo depois de abrir um `.docx` não grava sobre os bytes dele.
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
 * O documento nascido no editor salvo em `.docx`, sobre o pacote que o sidecar cria.
 * Depende do sidecar publicado, e não do corpus.
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

    // O original que segue é o pacote mínimo, e não os bytes gravados.
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

    // Mesmo pacote e mesmo conteúdo, mesmos bytes.
    expect(Buffer.compare(second.original, first.original)).toBe(0)
    expect(Buffer.compare(Buffer.from(second.bytes), Buffer.from(first.bytes))).toBe(0)
  })

  it('gravar cinco vezes o mesmo documento novo não acumula partes', async () => {
    // Partindo sempre do pacote mínimo, gravar de novo não soma numeração nem imagem
    // ao que já estava lá.
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

    // E não passa vazio: a lista e a imagem estão no pacote.
    expect(counted[0]).toEqual({ media: 1, numbering: 1 })
    expect(counted).toEqual(counted.map(() => counted[0]))
  })

  it('declara a perda do pacote de origem quando o modelo tem oid e o original não está aqui', async () => {
    // Um modelo com `oid` gravado sem o pacote dele perde estilos, notas e
    // comentários: a pessoa tem de ler isso na tela.
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
    // O `.sdoc` que foi `.docx` traz as faixas do arquivo de origem.
    const content = JSON.stringify({
      page: { ...page, headerBand: { left: [], center: [], right: [], rule: true, rows: [] } },
      doc: { type: 'doc', content: [] },
    })

    const saved = await saveDocx(client, content, { origin: null, destination: '/tmp/faixa.docx' })

    expect(saved.inventory.lost).toContain('cabeçalho e rodapé do arquivo .docx de origem')
  })
})

/**
 * O caminho que reconhece o pacote original vem do diálogo, na abertura, e do
 * `origin` do renderer, na gravação: os dois têm de bater, senão o `.docx` aberto é
 * gravado sobre o pacote mínimo.
 */
describe.skipIf(!published)('o caminho do pacote original', () => {
  let directory = ''

  beforeAll(async () => {
    directory = await mkdtemp(join(tmpdir(), 'librevia-caminho-'))
  })
  afterAll(() => rm(directory, { recursive: true, force: true }))

  /** Um `.docx` de verdade no disco, gravado pelo próprio escritor. */
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

  /** O mesmo arquivo escrito de outro jeito, sem `join`, que já normalizaria. */
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
    // A recuperação relê os bytes do disco pelo caminho do rascunho.
    const path = await docxOnDisk('recuperado.docx')
    forgetOpenedDocx()

    // O modelo vem do rascunho; aqui, de uma leitura logo esquecida.
    const opened = await openDocx(client, path)
    forgetOpenedDocx()

    expect(await adoptDocxOriginal(detoured(path))).toBe(true)
    const saved = await saveDocx(client, opened.content, { origin: path, destination: path })

    expect(saved.inventory.lost).toEqual([])
    expect(Buffer.compare(saved.original, await readFile(path))).toBe(0)
  })
})

/** As entradas do ZIP pelo diretório central, também as que têm descritor de dados. */
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
