import { describe, expect, it } from 'vitest'
import { DocumentKind } from '@shared/types.js'
import {
  buildWindowTitle,
  ensureSupportedExtension,
  extensionOf,
  fileNameFromPath,
  isOdtPath,
  isPdfPath,
  isReadableExtension,
  kindFromPath,
  isSupportedExtension,
  isWordPackagePath,
  isWordTemplatePath,
  OPENABLE_EXTENSIONS,
} from './formats.js'

describe('fileNameFromPath', () => {
  it.each([
    ['/home/ana/relatorio.txt', 'relatorio.txt'],
    ['C:\\Users\\Ana\\relatorio.txt', 'relatorio.txt'],
    ['\\\\servidor\\setor\\ata.txt', 'ata.txt'],
    ['/mnt/rede/contratos/minuta final.txt', 'minuta final.txt'],
    ['sem-pasta.txt', 'sem-pasta.txt'],
  ])('extracts the name from %s', (path, expected) => {
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
  ])('reads the extension of %s', (path, expected) => {
    expect(extensionOf(path)).toBe(expected)
  })

  it('does not mistake a dot in a folder name for the file extension', () => {
    expect(extensionOf('/home/ana/pasta.com.ponto/arquivo')).toBe('')
  })
})

describe('isSupportedExtension', () => {
  it.each(['/a/b.txt', '/a/b.sdoc', '/a/b.ssheet', '/a/b.docx', '/a/b.DOCX', '/a/b.xlsx', '/a/b.XLSX'])(
    'accepts %s',
    (path) => {
      expect(isSupportedExtension(path)).toBe(true)
    },
  )

  it.each(['/a/b.exe', '/a/b.sh', '/a/b', '/a/b.xls', '/a/b.ods'])('refuses %s', (path) => {
    // `.xls` and `.ods` are different formats, not variants: opening one as `.xlsx` would give a
    // corrupt file error.
    expect(isSupportedExtension(path)).toBe(false)
  })
})

describe('PDF', () => {
  it.each(['/a/b.pdf', '/a/b.PDF'])('%s opens in the viewer', (path) => {
    expect(isPdfPath(path)).toBe(true)
    expect(kindFromPath(path)).toBe(DocumentKind.Pdf)
  })

  it('is never a save target', () => {
    // Only viewed: saving text into a `.pdf` would produce a file no reader opens.
    expect(isSupportedExtension('/a/b.pdf')).toBe(false)
    expect(ensureSupportedExtension('/a/b.pdf')).toBe('/a/b.pdf.sdoc')
  })
})

describe('ODT', () => {
  it.each(['/a/b.odt', '/a/b.ODT'])('%s is read into a document', (path) => {
    expect(isOdtPath(path)).toBe(true)
    expect(isReadableExtension(path)).toBe(true)
    expect(kindFromPath(path)).toBe(DocumentKind.Document)
    expect(OPENABLE_EXTENSIONS).toContain('.odt')
  })

  it('is a save target, which main grants only to a document opened from an .odt', () => {
    expect(isSupportedExtension('/a/b.odt')).toBe(true)
    expect(isWordPackagePath('/a/b.odt')).toBe(false)
    expect(ensureSupportedExtension('/a/b.odt')).toBe('/a/b.odt')
  })

  it('a PDF is only viewed, not read into the editor', () => {
    expect(isReadableExtension('/a/b.pdf')).toBe(false)
  })
})

describe('Word templates (M11)', () => {
  it.each(['/a/b.dotx', '/a/b.DOTX', '/a/b.dotm'])('%s is a Word template and package', (path) => {
    expect(isSupportedExtension(path)).toBe(true)
    expect(isWordTemplatePath(path)).toBe(true)
    expect(isWordPackagePath(path)).toBe(true)
  })

  it('.docx is a Word package, but not a template', () => {
    expect(isWordTemplatePath('/a/b.docx')).toBe(false)
    expect(isWordPackagePath('/a/b.docx')).toBe(true)
  })

  it('.dotx is a save target; .dotm is not, because macros do not travel', () => {
    expect(ensureSupportedExtension('/a/b.dotx', DocumentKind.Document)).toBe('/a/b.dotx')
    expect(ensureSupportedExtension('/a/b.dotm', DocumentKind.Document)).toBe('/a/b.dotm.sdoc')
  })
})

describe('buildWindowTitle', () => {
  it('marks unsaved changes', () => {
    expect(buildWindowTitle('ata.txt', true, 'Librevia')).toBe('• ata.txt — Librevia')
  })

  it('does not mark when saved', () => {
    expect(buildWindowTitle('ata.txt', false, 'Librevia')).toBe('ata.txt — Librevia')
  })

  it('uses its own label for a file that has no name yet', () => {
    expect(buildWindowTitle(null, false, 'Librevia')).toBe('Sem título — Librevia')
  })
})
