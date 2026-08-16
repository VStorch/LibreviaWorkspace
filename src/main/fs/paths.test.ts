import { mkdir, mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import {
  MAX_FILE_BYTES,
  assertPathAuthorized,
  assertReadableFile,
  authorizePath,
  isPathAuthorized,
  resetAuthorizedPaths,
} from './paths.js'

let directory: string

beforeEach(async () => {
  directory = await mkdtemp(join(tmpdir(), 'librevia-paths-'))
  resetAuthorizedPaths()
})

afterEach(async () => {
  await rm(directory, { recursive: true, force: true })
})

describe('autorização de caminhos', () => {
  it('recusa um caminho que nunca foi escolhido pelo usuário', () => {
    // Este é o controle central: mesmo que o renderer seja comprometido por um
    // documento malicioso, ele não consegue mandar o main gravar em /etc.
    expect(() => assertPathAuthorized('/etc/passwd')).toThrow(/recusada/i)
  })

  it('aceita um caminho depois de autorizado', () => {
    const path = join(directory, 'ata.txt')
    authorizePath(path)

    expect(isPathAuthorized(path)).toBe(true)
    expect(() => assertPathAuthorized(path)).not.toThrow()
  })

  it('não se deixa enganar por caminho equivalente com ".."', () => {
    const path = join(directory, 'ata.txt')
    authorizePath(path)

    // O mesmo arquivo escrito de outro jeito continua autorizado…
    expect(isPathAuthorized(join(directory, 'sub', '..', 'ata.txt'))).toBe(true)
    // …e um vizinho alcançado por ".." continua recusado.
    expect(isPathAuthorized(join(directory, '..', 'outro.txt'))).toBe(false)
  })

  it('esquece as autorizações ao reiniciar a sessão', () => {
    authorizePath(join(directory, 'ata.txt'))
    resetAuthorizedPaths()

    expect(isPathAuthorized(join(directory, 'ata.txt'))).toBe(false)
  })
})

describe('assertReadableFile', () => {
  it('aceita um .txt comum', async () => {
    const path = join(directory, 'ok.txt')
    await writeFile(path, 'conteúdo')

    await expect(assertReadableFile(path)).resolves.toBeUndefined()
  })

  it('recusa caminho relativo', async () => {
    await expect(assertReadableFile('relativo.txt')).rejects.toMatchObject({
      code: 'INVALID_REQUEST',
    })
  })

  it('recusa extensão que o aplicativo não abre', async () => {
    // `.ods` é outro formato, não uma variação: abri-lo como `.xlsx` daria erro
    // de arquivo corrompido, que manda o usuário procurar o problema no lugar
    // errado.
    const path = join(directory, 'planilha.ods')
    await writeFile(path, 'x')

    await expect(assertReadableFile(path)).rejects.toMatchObject({ code: 'UNSUPPORTED_FORMAT' })
  })

  it('aceita .xlsx desde a Fase 7', async () => {
    const path = join(directory, 'vendas.xlsx')
    await writeFile(path, 'x')

    await expect(assertReadableFile(path)).resolves.toBeUndefined()
  })

  it('aceita .docx desde a Fase 4', async () => {
    const path = join(directory, 'contrato.docx')
    await writeFile(path, 'x')

    await expect(assertReadableFile(path)).resolves.toBeUndefined()
  })

  it('recusa uma pasta', async () => {
    const path = join(directory, 'pasta.txt')
    await mkdir(path)

    await expect(assertReadableFile(path)).rejects.toMatchObject({ code: 'NOT_A_FILE' })
  })

  it('recusa arquivo inexistente', async () => {
    await expect(assertReadableFile(join(directory, 'fantasma.txt'))).rejects.toMatchObject({
      code: 'FILE_NOT_FOUND',
    })
  })

  it('recusa arquivo acima do limite', async () => {
    const path = join(directory, 'enorme.txt')
    await writeFile(path, Buffer.alloc(MAX_FILE_BYTES + 1))

    await expect(assertReadableFile(path)).rejects.toMatchObject({ code: 'FILE_TOO_LARGE' })
  })
})
