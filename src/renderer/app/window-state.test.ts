import { describe, expect, it } from 'vitest'
import { DocumentKind, MenuCommand, RevisionView } from '@shared/types.js'
import { menuSnapshot } from './window-state.js'

const document = {
  file: { path: '/docs/a.docx', name: 'a.docx', kind: DocumentKind.Document },
  readOnly: false,
  workbook: null,
  trackChanges: true,
  revisionView: RevisionView.Simple,
}

describe('menu snapshot', () => {
  it('checks track changes and the revision view of the document on screen', () => {
    expect(menuSnapshot(document)).toEqual({
      activeKind: DocumentKind.Document,
      editable: true,
      checked: [MenuCommand.ToggleTrackChanges, MenuCommand.ShowSimpleMarkup],
      disabled: [],
    })
  })

  it('leaves track changes unchecked on a spreadsheet, which has none', () => {
    const sheet = menuSnapshot({
      ...document,
      file: { path: null, name: 'Planilha.xlsx', kind: DocumentKind.Spreadsheet },
      workbook: { sheets: [], activeSheet: 0 } as never,
    })

    expect(sheet.checked).not.toContain(MenuCommand.ToggleTrackChanges)
    expect(sheet.activeKind).toBe(DocumentKind.Spreadsheet)
  })

  it('says Home has no kind and nothing to edit', () => {
    const home = menuSnapshot({ ...document, file: null, trackChanges: undefined })

    expect(home.activeKind).toBeNull()
    expect(home.editable).toBe(false)
  })

  it('is not editable while the document is locked', () => {
    expect(menuSnapshot({ ...document, readOnly: true }).editable).toBe(false)
  })
})
