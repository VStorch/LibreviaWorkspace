import { afterEach, describe, expect, it, vi } from 'vitest'
import { DocumentKind, MenuCommand, RevisionView } from '@shared/types.js'
import { addTab, useTabs } from '../state/tabs.js'
import { createWorkspaceStore } from '../state/workspace-store.js'
import { menuSnapshot, watchEveryTab } from './window-state.js'

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

describe('watching the tabs', () => {
  const initialTabs = useTabs.getState()
  afterEach(() => useTabs.setState(initialTabs, true))

  it('hears a hidden tab change, as the quit question must know it is modified', () => {
    const hidden = createWorkspaceStore()
    addTab(hidden, null)
    addTab(createWorkspaceStore(), null)
    const changed = vi.fn()
    const stop = watchEveryTab(changed)

    hidden.setState({ isDirty: true })
    expect(changed).toHaveBeenCalledOnce()

    stop()
    hidden.setState({ isDirty: false })
    expect(changed).toHaveBeenCalledOnce()
  })

  it('hears a tab opened after it started', () => {
    const changed = vi.fn()
    const stop = watchEveryTab(changed)
    const later = createWorkspaceStore()
    addTab(later, null)
    changed.mockClear()

    later.setState({ isDirty: true })

    expect(changed).toHaveBeenCalledOnce()
    stop()
  })
})
