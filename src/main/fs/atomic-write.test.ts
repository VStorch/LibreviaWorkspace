import { chmod, mkdtemp, readFile, readdir, rm, stat, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { AppError } from '@shared/errors.js'
import { writeFileAtomic } from './atomic-write.js'

/**
 * Three tests only hold on POSIX: on Windows Node's `chmod` only sets read-only, `stat` returns
 * 0o666 and a directory stays writable. The protection they cover is covered on Linux.
 */
const emPosix = process.platform !== 'win32'

let directory: string

beforeEach(async () => {
  directory = await mkdtemp(join(tmpdir(), 'librevia-test-'))
})

afterEach(async () => {
  await chmod(directory, 0o700).catch(() => undefined)
  await rm(directory, { recursive: true, force: true })
})

describe('writeFileAtomic', () => {
  it('cria um arquivo novo', async () => {
    const target = join(directory, 'novo.txt')
    await writeFileAtomic(target, 'conteúdo inicial')

    expect(await readFile(target, 'utf8')).toBe('conteúdo inicial')
  })

  it('não deixa arquivo temporário para trás', async () => {
    await writeFileAtomic(join(directory, 'a.txt'), 'x')

    // A forgotten .tmp in the user's folder is visible litter, and on a shared network folder
    // everyone sees it.
    const leftovers = (await readdir(directory)).filter((name) => name.endsWith('.tmp'))
    expect(leftovers).toEqual([])
  })

  it('guarda o conteúdo anterior em .bak ao sobrescrever', async () => {
    const target = join(directory, 'ata.txt')
    await writeFileAtomic(target, 'versão 1')
    await writeFileAtomic(target, 'versão 2')

    expect(await readFile(target, 'utf8')).toBe('versão 2')
    // The .bak must hold the *previous* version: that is the protection against a mistaken save
    // over a good document.
    expect(await readFile(`${target}.bak`, 'utf8')).toBe('versão 1')
  })

  it('não cria .bak quando o arquivo ainda não existia', async () => {
    const target = join(directory, 'primeiro.txt')
    await writeFileAtomic(target, 'conteúdo')

    await expect(stat(`${target}.bak`)).rejects.toThrow()
  })

  it.runIf(emPosix)('preserva as permissões do arquivo existente', async () => {
    const target = join(directory, 'compartilhado.txt')
    await writeFile(target, 'original')
    await chmod(target, 0o640)

    await writeFileAtomic(target, 'atualizado')

    // Saving must not narrow access to a file shared by a team.
    expect((await stat(target)).mode & 0o777).toBe(0o640)
  })

  it.runIf(emPosix)('mantém o arquivo original intacto quando a gravação falha', async () => {
    const target = join(directory, 'protegido.txt')
    await writeFile(target, 'conteúdo valioso')
    await chmod(directory, 0o500) // read and traverse, no write

    await expect(writeFileAtomic(target, 'tentativa')).rejects.toBeInstanceOf(AppError)

    await chmod(directory, 0o700)
    expect(await readFile(target, 'utf8')).toBe('conteúdo valioso')
  })

  it.runIf(emPosix)('reporta falha de gravação com mensagem compreensível', async () => {
    await chmod(directory, 0o500)

    await expect(writeFileAtomic(join(directory, 'x.txt'), 'a')).rejects.toMatchObject({
      message: expect.stringMatching(/permissão|salvar/i),
    })
  })

  it('grava num sistema de arquivos diferente do temporário do sistema', async () => {
    // The network folder test. An implementation that prepared the temp file in os.tmpdir() and
    // then renamed it would fail here with EXDEV, which is exactly what happens on a mounted share.
    // It passes only because the temp file is born in the destination folder.
    const crossDevice = await mkdtemp(join('/dev/shm', 'librevia-xdev-')).catch(() => null)
    if (crossDevice === null) return // no separate tmpfs here

    try {
      const target = join(crossDevice, 'em-outro-volume.txt')
      await writeFileAtomic(target, 'primeira versão')
      await writeFileAtomic(target, 'segunda versão')

      expect(await readFile(target, 'utf8')).toBe('segunda versão')
      expect(await readFile(`${target}.bak`, 'utf8')).toBe('primeira versão')
    } finally {
      await rm(crossDevice, { recursive: true, force: true })
    }
  })

  it('grava conteúdo com acentuação e quebras de linha sem alterar bytes', async () => {
    const target = join(directory, 'acentos.txt')
    const content = 'Ação\nCoração — “aspas”\r\nfim\t.'
    await writeFileAtomic(target, content)

    expect(await readFile(target, 'utf8')).toBe(content)
  })
})
