import { describe, expect, it } from 'vitest'
import { DocumentKind, SaveFormat } from '@shared/types.js'
import {
  defaultFileName,
  filtersStartingWith,
  preferredExtension,
  withExtension,
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

describe('preferred extension', () => {
  it('writes for Word and Excel when the user follows Office or Google', () => {
    expect(preferredExtension(DocumentKind.Document, SaveFormat.Office)).toBe('.docx')
    expect(preferredExtension(DocumentKind.Spreadsheet, SaveFormat.Office)).toBe('.xlsx')
  })

  it('writes .odt for LibreOffice, and .xlsx for its spreadsheets, since .ods is not written yet', () => {
    expect(preferredExtension(DocumentKind.Document, SaveFormat.LibreOffice)).toBe('.odt')
    expect(preferredExtension(DocumentKind.Spreadsheet, SaveFormat.LibreOffice)).toBe('.xlsx')
  })

  it('keeps the own formats while the user has not chosen', () => {
    expect(preferredExtension(DocumentKind.Document, SaveFormat.Ask)).toBe('.sdoc')
    expect(preferredExtension(DocumentKind.Spreadsheet, SaveFormat.Librevia)).toBe('.ssheet')
  })

  it('names a new file with it', () => {
    expect(defaultFileName(DocumentKind.Document, undefined, SaveFormat.Office)).toBe(
      'Documento sem título.docx',
    )
  })
})

describe('withExtension', () => {
  it('swaps the last extension only', () => {
    expect(withExtension('relatorio.v2.sdoc', '.docx')).toBe('relatorio.v2.docx')
  })

  it('adds one to a name without it, and keeps a leading dot as part of the name', () => {
    expect(withExtension('ata', '.odt')).toBe('ata.odt')
    expect(withExtension('.config', '.odt')).toBe('.config.odt')
  })
})

describe('ensureSupportedExtension with the suggested extension', () => {
  it('completes a name typed without extension with the one the dialog suggested', () => {
    expect(ensureSupportedExtension('/tmp/ata', DocumentKind.Document, '.docx')).toBe('/tmp/ata.docx')
  })
})

describe('filtersStartingWith', () => {
  const filters = [
    { name: 'Librevia', extensions: ['sdoc'] },
    { name: 'Word', extensions: ['docx'] },
    { name: 'ODT', extensions: ['odt'] },
  ]

  it('puts first the filter of the suggested name, so Save as of an .odt starts on ODT', () => {
    expect(filtersStartingWith(filters, 'ata.odt').map((filter) => filter.name)).toEqual([
      'ODT',
      'Librevia',
      'Word',
    ])
  })

  it('keeps the order when no filter matches', () => {
    expect(filtersStartingWith(filters, 'ata').map((filter) => filter.name)).toEqual([
      'Librevia',
      'Word',
      'ODT',
    ])
  })
})
