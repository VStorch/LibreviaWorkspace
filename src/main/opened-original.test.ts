import { beforeEach, describe, expect, it } from 'vitest'
import {
  followOriginal,
  isCreatedOriginal,
  originalFormatOf,
  originalOf,
  OriginalFormat,
  releaseOriginal,
  rememberOriginal,
  resetOriginals,
} from './opened-original.js'

const first = '/home/ana/relatorio.docx'
const second = '/home/ana/ata.odt'
const sheet = '/home/ana/contas.xlsx'

beforeEach(() => {
  resetOriginals()
})

describe('originals by path', () => {
  it('keeps the original of every opened file', () => {
    rememberOriginal(first, Buffer.from('first'), OriginalFormat.Docx)
    rememberOriginal(second, Buffer.from('second'), OriginalFormat.Odt, true)
    rememberOriginal(sheet, Buffer.from('sheet'), OriginalFormat.Xlsx)

    expect(originalOf(first, OriginalFormat.Docx)?.toString()).toBe('first')
    expect(originalOf(second, OriginalFormat.Odt)?.toString()).toBe('second')
    expect(originalOf(sheet, OriginalFormat.Xlsx)?.toString()).toBe('sheet')
    expect(isCreatedOriginal(second)).toBe(true)
    expect(isCreatedOriginal(first)).toBe(false)
  })

  it('reads by normalized path and never lends another format', () => {
    rememberOriginal(first, Buffer.from('first'), OriginalFormat.Docx)

    expect(originalOf('/home/ana/../ana/relatorio.docx', OriginalFormat.Docx)).not.toBeNull()
    expect(originalOf(first, OriginalFormat.Odt)).toBeNull()
    expect(originalFormatOf(first)).toBe(OriginalFormat.Docx)
    expect(originalFormatOf(null)).toBeNull()
    expect(originalFormatOf(second)).toBeNull()
  })

  it('releases only the closed file', () => {
    rememberOriginal(first, Buffer.from('first'), OriginalFormat.Docx)
    rememberOriginal(sheet, Buffer.from('sheet'), OriginalFormat.Xlsx)

    releaseOriginal(first)

    expect(originalOf(first, OriginalFormat.Docx)).toBeNull()
    expect(originalOf(sheet, OriginalFormat.Xlsx)).not.toBeNull()
  })

  it('moves the original to the destination of a save as', () => {
    rememberOriginal(first, Buffer.from('first'), OriginalFormat.Docx)
    rememberOriginal(sheet, Buffer.from('sheet'), OriginalFormat.Xlsx)
    const destination = '/home/ana/copia.xlsx'

    followOriginal(sheet, destination, null)

    expect(originalOf(sheet, OriginalFormat.Xlsx)).toBeNull()
    expect(originalOf(destination, OriginalFormat.Xlsx)?.toString()).toBe('sheet')
    expect(originalOf(first, OriginalFormat.Docx)).not.toBeNull()
  })

  it('replaces the moved original with the package the save started from', () => {
    rememberOriginal(first, Buffer.from('first'), OriginalFormat.Docx)
    const destination = '/home/ana/copia.odt'

    followOriginal(first, destination, {
      original: Buffer.from('minimal'),
      format: OriginalFormat.Odt,
      created: true,
    })

    expect(originalOf(first, OriginalFormat.Docx)).toBeNull()
    expect(originalOf(destination, OriginalFormat.Odt)?.toString()).toBe('minimal')
    expect(isCreatedOriginal(destination)).toBe(true)
  })

  it('keeps the original when saving over the same file', () => {
    rememberOriginal(first, Buffer.from('first'), OriginalFormat.Docx)

    followOriginal(first, first, { original: Buffer.from('first'), format: OriginalFormat.Docx })

    expect(originalOf(first, OriginalFormat.Docx)?.toString()).toBe('first')
  })
})
