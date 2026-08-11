import { mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { readTextFile } from './read-text.js'

let directory: string

beforeEach(async () => {
  directory = await mkdtemp(join(tmpdir(), 'librevia-read-'))
})

afterEach(async () => {
  await rm(directory, { recursive: true, force: true })
})

async function writeBytes(name: string, bytes: Buffer): Promise<string> {
  const path = join(directory, name)
  await writeFile(path, bytes)
  return path
}

describe('readTextFile', () => {
  it('lê UTF-8 sem BOM', async () => {
    const path = await writeBytes('simples.txt', Buffer.from('Relatório de março', 'utf8'))
    expect(await readTextFile(path)).toBe('Relatório de março')
  })

  it('remove o BOM de UTF-8', async () => {
    // É o que o Bloco de Notas do Windows grava — e o BOM apareceria como um
    // caractere invisível no começo do documento.
    const bytes = Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from('Ata', 'utf8')])
    const path = await writeBytes('bom-utf8.txt', bytes)

    expect(await readTextFile(path)).toBe('Ata')
  })

  it('lê UTF-16 little endian', async () => {
    const bytes = Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from('Contrato', 'utf16le')])
    const path = await writeBytes('utf16le.txt', bytes)

    expect(await readTextFile(path)).toBe('Contrato')
  })

  it('lê UTF-16 big endian', async () => {
    const swapped = Buffer.from('Contrato', 'utf16le')
    swapped.swap16()
    const path = await writeBytes('utf16be.txt', Buffer.concat([Buffer.from([0xfe, 0xff]), swapped]))

    expect(await readTextFile(path)).toBe('Contrato')
  })

  it('recusa conteúdo binário', async () => {
    // Abrir um executável por engano encheria o editor de lixo — e salvá-lo
    // destruiria o arquivo.
    const path = await writeBytes('binario.txt', Buffer.from([0x7f, 0x45, 0x4c, 0x46, 0x00, 0x01]))

    await expect(readTextFile(path)).rejects.toMatchObject({ code: 'NOT_TEXT_FILE' })
  })

  it('aceita arquivo vazio', async () => {
    const path = await writeBytes('vazio.txt', Buffer.alloc(0))
    expect(await readTextFile(path)).toBe('')
  })

  it('reporta arquivo inexistente com mensagem compreensível', async () => {
    await expect(readTextFile(join(directory, 'nao-existe.txt'))).rejects.toMatchObject({
      code: 'FILE_NOT_FOUND',
    })
  })
})
