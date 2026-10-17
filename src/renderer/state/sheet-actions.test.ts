import { describe, expect, it } from 'vitest'
import { createStore } from 'zustand/vanilla'
import { writeText } from '@services/spreadsheet/edit.js'
import { createSheet, getCell, type WorkbookModel } from '@services/spreadsheet/model.js'
import { createSheetActions } from './sheet-actions.js'
import type { WorkspaceState } from './types.js'

type Fields = Pick<WorkspaceState, 'workbook' | 'generation' | 'isDirty' | 'properties'>

function openWorkbook() {
  const store = createStore<WorkspaceState>()(
    (set, get, api) =>
      ({
        workbook: null,
        generation: 0,
        isDirty: false,
        properties: undefined,
        ...createSheetActions(set, get, api.subscribe),
      }) as Fields as WorkspaceState,
  )
  const workbook: WorkbookModel = { sheets: [createSheet('Planilha1')], activeSheet: 0 }
  store.setState({ workbook, generation: 1, isDirty: false })
  return store
}

const valueAt = (state: WorkspaceState, ref: { row: number; column: number }): unknown =>
  getCell(state.workbook!.sheets[0]!, ref.row, ref.column)?.value

function write(store: ReturnType<typeof openWorkbook>, column: number, text: string): void {
  const state = store.getState()
  state.updateSheet(writeText(state.workbook!.sheets[0]!, 0, column, text))
}

describe('desfazer da planilha', () => {
  it('desfazer e refazer não gravam passo novo; uma edição nova apaga o refazer', () => {
    const store = openWorkbook()
    write(store, 0, 'a')
    write(store, 1, 'b')
    store.getState().undoSheet()
    expect(valueAt(store.getState(), { row: 0, column: 1 })).toBeUndefined()
    store.getState().redoSheet()
    expect(valueAt(store.getState(), { row: 0, column: 1 })).toBe('b')
    store.getState().undoSheet()
    store.getState().undoSheet()
    expect(valueAt(store.getState(), { row: 0, column: 0 })).toBeUndefined()

    write(store, 2, 'c')
    store.getState().redoSheet()
    expect(valueAt(store.getState(), { row: 0, column: 0 })).toBeUndefined()
  })

  it('a mesma planilha de volta não conta como edição', () => {
    const store = openWorkbook()
    const state = store.getState()
    state.updateSheet(state.workbook!.sheets[0]!)
    expect(store.getState().isDirty).toBe(false)
    store.getState().undoSheet()
    expect(store.getState().workbook).toBe(state.workbook)
  })

  it('desfazer até o estado salvo deixa de pedir para salvar, e refazer volta a pedir', () => {
    const store = openWorkbook()
    write(store, 0, 'a')
    expect(store.getState().isDirty).toBe(true)
    store.getState().undoSheet()
    expect(store.getState().isDirty).toBe(false)
    store.getState().redoSheet()
    expect(store.getState().isDirty).toBe(true)

    store.setState({ isDirty: false })
    write(store, 1, 'b')
    store.getState().undoSheet()
    expect(store.getState().isDirty).toBe(false)
    store.getState().undoSheet()
    expect(store.getState().isDirty).toBe(true)
  })

  it('abrir outro arquivo esquece o histórico do anterior', () => {
    const store = openWorkbook()
    write(store, 0, 'a')
    const other: WorkbookModel = { sheets: [createSheet('Outra')], activeSheet: 0 }
    store.setState({ workbook: other, generation: 2, isDirty: false })
    store.getState().undoSheet()
    expect(store.getState().workbook).toBe(other)
  })
})
