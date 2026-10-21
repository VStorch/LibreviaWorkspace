import { mkdir, mkdtemp, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import {
  MAX_FILE_BYTES,
  assertPathAuthorized,
  assertReadableFile,
  assertViewablePdf,
  allowViewing,
  assertViewingAllowed,
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

describe('path authorization', () => {
  it('refuses a path the user never chose', () => {
    // The central control: even if a malicious document compromises the renderer, it cannot make
    // main write to /etc.
    expect(() => assertPathAuthorized('/etc/passwd')).toThrow(/recusada/i)
  })

  it('accepts a path once authorized', () => {
    const path = join(directory, 'ata.txt')
    authorizePath(path)

    expect(isPathAuthorized(path)).toBe(true)
    expect(() => assertPathAuthorized(path)).not.toThrow()
  })

  it('is not fooled by an equivalent path with ".."', () => {
    const path = join(directory, 'ata.txt')
    authorizePath(path)

    // The same file spelled differently stays authorized…
    expect(isPathAuthorized(join(directory, 'sub', '..', 'ata.txt'))).toBe(true)
    // …and a neighbor reached through ".." stays refused.
    expect(isPathAuthorized(join(directory, '..', 'outro.txt'))).toBe(false)
  })

  it('forgets the authorizations when the session restarts', () => {
    authorizePath(join(directory, 'ata.txt'))
    resetAuthorizedPaths()

    expect(isPathAuthorized(join(directory, 'ata.txt'))).toBe(false)
  })
})

describe('assertReadableFile', () => {
  it('accepts a plain .txt', async () => {
    const path = join(directory, 'ok.txt')
    await writeFile(path, 'content')

    await expect(assertReadableFile(path)).resolves.toBeUndefined()
  })

  it('refuses a relative path', async () => {
    await expect(assertReadableFile('relativo.txt')).rejects.toMatchObject({
      code: 'INVALID_REQUEST',
    })
  })

  it('refuses an extension the app does not open', async () => {
    // `.ods` is another format, not a variant: opening it as `.xlsx` would give a corrupt file
    // error, sending the user to look for the problem in the wrong place.
    const path = join(directory, 'planilha.ods')
    await writeFile(path, 'x')

    await expect(assertReadableFile(path)).rejects.toMatchObject({ code: 'UNSUPPORTED_FORMAT' })
  })

  it('accepts .xlsx', async () => {
    const path = join(directory, 'vendas.xlsx')
    await writeFile(path, 'x')

    await expect(assertReadableFile(path)).resolves.toBeUndefined()
  })

  it('accepts .docx', async () => {
    const path = join(directory, 'contrato.docx')
    await writeFile(path, 'x')

    await expect(assertReadableFile(path)).resolves.toBeUndefined()
  })

  it('refuses a folder', async () => {
    const path = join(directory, 'pasta.txt')
    await mkdir(path)

    await expect(assertReadableFile(path)).rejects.toMatchObject({ code: 'NOT_A_FILE' })
  })

  it('refuses a missing file', async () => {
    await expect(assertReadableFile(join(directory, 'fantasma.txt'))).rejects.toMatchObject({
      code: 'FILE_NOT_FOUND',
    })
  })

  it('refuses a file above the limit', async () => {
    const path = join(directory, 'enorme.txt')
    await writeFile(path, Buffer.alloc(MAX_FILE_BYTES + 1))

    await expect(assertReadableFile(path)).rejects.toMatchObject({ code: 'FILE_TOO_LARGE' })
  })
})

describe('assertViewablePdf', () => {
  it('accepts a PDF above the editable files limit', async () => {
    // Chromium streams the PDF: its bytes never cross to the renderer, so the limit does not apply.
    const path = join(directory, 'digitalizado.pdf')
    await writeFile(path, Buffer.alloc(MAX_FILE_BYTES + 1))

    await expect(assertViewablePdf(path)).resolves.toBeUndefined()
  })

  it('refuses what is not a PDF', async () => {
    const path = join(directory, 'ata.txt')
    await writeFile(path, 'ata')

    await expect(assertViewablePdf(path)).rejects.toMatchObject({ code: 'UNSUPPORTED_FORMAT' })
  })

  it('refuses a missing PDF', async () => {
    await expect(assertViewablePdf(join(directory, 'fantasma.pdf'))).rejects.toMatchObject({
      code: 'FILE_NOT_FOUND',
    })
  })
})

describe('view-only PDFs', () => {
  it('the viewer only shows a PDF that main opened', () => {
    // The renderer asks to show a path: one it made up does not reach the viewer.
    expect(() => assertViewingAllowed('/home/ana/outro.pdf')).toThrow(/recusada/i)

    const path = join(directory, 'relatorio.pdf')
    allowViewing(path)
    expect(assertViewingAllowed(path)).toBe(path)
  })

  it('viewing a PDF does not authorize writing to it', () => {
    const path = join(directory, 'relatorio.pdf')
    allowViewing(path)

    expect(isPathAuthorized(path)).toBe(false)
  })
})
