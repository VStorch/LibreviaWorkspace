import { describe, expect, it } from 'vitest'
import { DocumentKind } from '@shared/types.js'
import {
  buildWindowTitle,
  ensureSupportedExtension,
  extensionOf,
  fileNameFromPath,
  isInspectablePackagePath,
  isOdtPackagePath,
  isOdtPath,
  isOdtTemplatePath,
  isOfficePackagePath,
  isPdfPath,
  isOpenableExtension,
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
    expect(isSupportedExtension(path)).toBe(true)
    expect(kindFromPath(path)).toBe(DocumentKind.Document)
    expect(OPENABLE_EXTENSIONS).toContain('.odt')
  })

  it('is a save target for any document', () => {
    expect(isSupportedExtension('/a/b.odt')).toBe(true)
    expect(isWordPackagePath('/a/b.odt')).toBe(false)
    expect(ensureSupportedExtension('/a/b.odt')).toBe('/a/b.odt')
  })

  it('a PDF is only viewed, not read into the editor', () => {
    expect(isSupportedExtension('/a/b.pdf')).toBe(false)
  })

  it.each(['/a/b.odt', '/a/b.PDF', '/a/b.xlsx', '/a/b.dotm'])('%s can be opened', (path) => {
    expect(isOpenableExtension(path)).toBe(true)
  })

  it.each(['/a/b.exe', '/a/odt', '/a/b.html'])('%s cannot be opened', (path) => {
    expect(isOpenableExtension(path)).toBe(false)
  })
})

describe('ODT templates', () => {
  it.each(['/a/b.ott', '/a/b.OTT'])('%s is an ODT template and package, and a save target', (path) => {
    expect(isOdtTemplatePath(path)).toBe(true)
    expect(isOdtPackagePath(path)).toBe(true)
    expect(isOdtPath(path)).toBe(false)
    expect(isSupportedExtension(path)).toBe(true)
    expect(ensureSupportedExtension(path)).toBe(path)
    expect(kindFromPath(path)).toBe(DocumentKind.Document)
  })

  it('.odt is an ODT package, but not a template', () => {
    expect(isOdtTemplatePath('/a/b.odt')).toBe(false)
    expect(isOdtPackagePath('/a/b.odt')).toBe(true)
  })

  it.each(['/a/b.odt', '/a/b.ott', '/a/b.docx', '/a/b.dotx', '/a/b.xlsx'])(
    '%s goes through the personal data inspector',
    (path) => {
      expect(isInspectablePackagePath(path)).toBe(true)
    },
  )

  it('ODF packages stay out of the Office packages', () => {
    expect(isOfficePackagePath('/a/b.odt')).toBe(false)
    expect(isOfficePackagePath('/a/b.ott')).toBe(false)
    expect(isInspectablePackagePath('/a/b.sdoc')).toBe(false)
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
