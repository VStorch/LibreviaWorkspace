import { describe, expect, it } from 'vitest'
import { DocumentKind } from '@shared/types.js'
import {
  buildWindowTitle,
  ensureSupportedExtension,
  extensionOf,
  fileNameFromPath,
  isSupportedExtension,
  isWordPackagePath,
  isWordTemplatePath,
} from './formats.js'

describe('fileNameFromPath', () => {
  it.each([
    ['/home/ana/relatorio.txt', 'relatorio.txt'],
    ['C:\\Users\\Ana\\relatorio.txt', 'relatorio.txt'],
    ['\\\\servidor\\setor\\ata.txt', 'ata.txt'],
    ['/mnt/rede/contratos/minuta final.txt', 'minuta final.txt'],
    ['sem-pasta.txt', 'sem-pasta.txt'],
  ])('extrai o nome de %s', (path, expected) => {
    expect(fileNameFromPath(path)).toBe(expected)
  })
})

describe('extensionOf', () => {
  it.each([
    ['/a/b/c.txt', '.txt'],
    ['/a/b/c.TXT', '.txt'],
    ['/a/b/arquivo.com.ponto.txt', '.txt'],
    ['/a/b/sem-extensao', ''],
    // A leading dot is a hidden file, not an extension.
    ['/a/b/.oculto', ''],
  ])('lê a extensão de %s', (path, expected) => {
    expect(extensionOf(path)).toBe(expected)
  })

  it('não confunde ponto de pasta com extensão do arquivo', () => {
    expect(extensionOf('/home/ana/pasta.com.ponto/arquivo')).toBe('')
  })
})

describe('isSupportedExtension', () => {
  it.each(['/a/b.txt', '/a/b.sdoc', '/a/b.ssheet', '/a/b.docx', '/a/b.DOCX', '/a/b.xlsx', '/a/b.XLSX'])(
    'aceita %s',
    (path) => {
      expect(isSupportedExtension(path)).toBe(true)
    },
  )

  it.each(['/a/b.exe', '/a/b.sh', '/a/b', '/a/b.xls', '/a/b.ods'])('recusa %s', (path) => {
    // `.xls` and `.ods` are different formats, not variants: opening one as `.xlsx` would give a
    // corrupt file error.
    expect(isSupportedExtension(path)).toBe(false)
  })
})

describe('modelos do Word (M11)', () => {
  it.each(['/a/b.dotx', '/a/b.DOTX', '/a/b.dotm'])('%s é modelo e pacote do Word', (path) => {
    expect(isSupportedExtension(path)).toBe(true)
    expect(isWordTemplatePath(path)).toBe(true)
    expect(isWordPackagePath(path)).toBe(true)
  })

  it('o .docx é pacote do Word, mas não modelo', () => {
    expect(isWordTemplatePath('/a/b.docx')).toBe(false)
    expect(isWordPackagePath('/a/b.docx')).toBe(true)
  })

  it('o .dotx é destino; o .dotm não, porque as macros não viajam', () => {
    expect(ensureSupportedExtension('/a/b.dotx', DocumentKind.Document)).toBe('/a/b.dotx')
    expect(ensureSupportedExtension('/a/b.dotm', DocumentKind.Document)).toBe('/a/b.dotm.sdoc')
  })
})

describe('buildWindowTitle', () => {
  it('marca alterações não salvas', () => {
    expect(buildWindowTitle('ata.txt', true, 'Librevia')).toBe('• ata.txt — Librevia')
  })

  it('não marca quando está salvo', () => {
    expect(buildWindowTitle('ata.txt', false, 'Librevia')).toBe('ata.txt — Librevia')
  })

  it('usa rótulo próprio para arquivo ainda sem nome', () => {
    expect(buildWindowTitle(null, false, 'Librevia')).toBe('Sem título — Librevia')
  })
})
