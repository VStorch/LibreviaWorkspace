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
const onPosix = process.platform !== 'win32'

let directory: string

beforeEach(async () => {
  directory = await mkdtemp(join(tmpdir(), 'librevia-test-'))
})

afterEach(async () => {
  await chmod(directory, 0o700).catch(() => undefined)
  await rm(directory, { recursive: true, force: true })
})

describe('writeFileAtomic', () => {
  it('creates a new file', async () => {
    const target = join(directory, 'novo.txt')
    await writeFileAtomic(target, 'initial content')

    expect(await readFile(target, 'utf8')).toBe('initial content')
  })

  it('leaves no temporary file behind', async () => {
    await writeFileAtomic(join(directory, 'a.txt'), 'x')

    // A forgotten .tmp in the user's folder is visible litter, and on a shared network folder
    // everyone sees it.
    const leftovers = (await readdir(directory)).filter((name) => name.endsWith('.tmp'))
    expect(leftovers).toEqual([])
  })

  it('keeps the previous content in .bak when overwriting', async () => {
    const target = join(directory, 'ata.txt')
    await writeFileAtomic(target, 'version 1')
    await writeFileAtomic(target, 'version 2')

    expect(await readFile(target, 'utf8')).toBe('version 2')
    // The .bak must hold the *previous* version: that is the protection against a mistaken save
    // over a good document.
    expect(await readFile(`${target}.bak`, 'utf8')).toBe('version 1')
  })

  it('creates no .bak when the file did not exist yet', async () => {
    const target = join(directory, 'primeiro.txt')
    await writeFileAtomic(target, 'content')

    await expect(stat(`${target}.bak`)).rejects.toThrow()
  })

  it.runIf(onPosix)('keeps the permissions of the existing file', async () => {
    const target = join(directory, 'compartilhado.txt')
    await writeFile(target, 'original')
    await chmod(target, 0o640)

    await writeFileAtomic(target, 'atualizado')

    // Saving must not narrow access to a file shared by a team.
    expect((await stat(target)).mode & 0o777).toBe(0o640)
  })

  it.runIf(onPosix)('keeps the original file intact when the write fails', async () => {
    const target = join(directory, 'protegido.txt')
    await writeFile(target, 'valuable content')
    await chmod(directory, 0o500) // read and traverse, no write

    await expect(writeFileAtomic(target, 'tentativa')).rejects.toBeInstanceOf(AppError)

    await chmod(directory, 0o700)
    expect(await readFile(target, 'utf8')).toBe('valuable content')
  })

  it.runIf(onPosix)('reports a write failure with an understandable message', async () => {
    await chmod(directory, 0o500)

    await expect(writeFileAtomic(join(directory, 'x.txt'), 'a')).rejects.toMatchObject({
      message: expect.stringMatching(/permissão|salvar/i),
    })
  })

  it('writes to a file system different from the system temp one', async () => {
    // The network folder test. An implementation that prepared the temp file in os.tmpdir() and
    // then renamed it would fail here with EXDEV, which is exactly what happens on a mounted share.
    // It passes only because the temp file is born in the destination folder.
    const crossDevice = await mkdtemp(join('/dev/shm', 'librevia-xdev-')).catch(() => null)
    if (crossDevice === null) return // no separate tmpfs here

    try {
      const target = join(crossDevice, 'em-outro-volume.txt')
      await writeFileAtomic(target, 'first version')
      await writeFileAtomic(target, 'second version')

      expect(await readFile(target, 'utf8')).toBe('second version')
      expect(await readFile(`${target}.bak`, 'utf8')).toBe('first version')
    } finally {
      await rm(crossDevice, { recursive: true, force: true })
    }
  })

  it('writes accented content and line breaks without changing bytes', async () => {
    const target = join(directory, 'acentos.txt')
    const content = 'Ação\nCoração — “aspas”\r\nfim\t.'
    await writeFileAtomic(target, content)

    expect(await readFile(target, 'utf8')).toBe(content)
  })
})
