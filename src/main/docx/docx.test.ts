/**
 * O caminho que o processo main percorre de verdade ao abrir e salvar `.docx`.
 *
 * Roda contra o **corpus real**, apontado por `LIBREVIA_CORPUS_DIR`, e é pulado
 * quando a variável não existe — os arquivos têm marca de cliente e capturas de
 * sistemas internos, então não entram no repositório (docs/01-corpus-docx.md).
 * O CI cobre as mesmas estruturas com fixtures sintéticos, do lado C#.
 */

import { readdir, readFile } from 'node:fs/promises'
import { join } from 'node:path'
import { afterAll, describe, expect, it } from 'vitest'
import { SidecarClient } from '../sidecar/client.js'
import { sidecarPathIn } from '../sidecar/locate.js'
import { forgetOpenedDocx, openDocx, saveDocx } from './index.js'

const corpusDirectory = process.env['LIBREVIA_CORPUS_DIR']

const documents =
  corpusDirectory === undefined
    ? []
    : (await readdir(corpusDirectory))
        .filter((name) => name.toLowerCase().endsWith('.docx'))
        .map((name) => join(corpusDirectory, name))

const client = new SidecarClient(() => Promise.resolve(sidecarPathIn(process.cwd())))
afterAll(() => client.dispose())

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
    // Abrir e salvar por reflexo é o caso mais comum de todos, e precisa custar
    // zero: cada bloco reescrito é uma chance de perder o que não entendemos.
    forgetOpenedDocx()
    const opened = await openDocx(client, path)
    const saved = await saveDocx(client, opened.content)

    expect(saved.bytes.length).toBeGreaterThan(0)
    expect(saved.inventory.lost).toEqual([])
  })

  it('preserva todas as partes do pacote menos o corpo', async () => {
    const path = documents[0]!
    forgetOpenedDocx()

    const original = await readFile(path)
    const opened = await openDocx(client, path)
    const saved = await saveDocx(client, opened.content)

    const before = await listParts(original)
    const after = await listParts(Buffer.from(saved.bytes))

    expect([...after.keys()].sort()).toEqual([...before.keys()].sort())

    const changed = [...before.keys()].filter((name) => !before.get(name)!.equals(after.get(name)!))
    expect(changed).toEqual(['word/document.xml'])
  })

  it('recusa gravar .docx sem um original aberto', async () => {
    // Salvar como .docx um documento criado do zero não tem pacote para
    // preservar. Melhor recusar com explicação do que gerar um arquivo pobre.
    forgetOpenedDocx()

    await expect(saveDocx(client, JSON.stringify({ page: {}, doc: { type: 'doc' } }))).rejects.toThrow(
      /aberto a partir de um arquivo/i,
    )
  })
})

/** Lê as entradas de um ZIP sem depender de biblioteca. */
async function listParts(zip: Buffer): Promise<Map<string, Buffer>> {
  const { execFile } = await import('node:child_process')
  const { promisify } = await import('node:util')
  const { mkdtemp, writeFile, rm } = await import('node:fs/promises')
  const { tmpdir } = await import('node:os')

  const run = promisify(execFile)
  const directory = await mkdtemp(join(tmpdir(), 'librevia-zip-'))
  const archive = join(directory, 'a.zip')

  try {
    await writeFile(archive, zip)
    const { stdout } = await run('unzip', ['-Z1', archive])
    const names = stdout.split('\n').filter((line) => line.length > 0 && !line.endsWith('/'))

    const parts = new Map<string, Buffer>()
    for (const name of names) {
      // O unzip trata `[` e `]` como curinga, e o OOXML tem uma parte chamada
      // `[Content_Types].xml` — sem escapar, ela nunca é encontrada.
      const pattern = name.replace(/[[\]*?]/g, (char) => `\\${char}`)
      const { stdout: content } = await run('unzip', ['-p', archive, pattern], {
        encoding: 'buffer',
        maxBuffer: 64 * 1024 * 1024,
      })
      parts.set(name, Buffer.from(content))
    }
    return parts
  } finally {
    await rm(directory, { recursive: true, force: true })
  }
}
