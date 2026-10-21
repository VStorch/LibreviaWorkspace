import { describe, expect, it } from 'vitest'
import {
  changedCells,
  emptyHistory,
  MAX_UNDO_CELLS,
  MAX_UNDO_STEPS,
  recordEdit,
  redoStep,
  undoStep,
  type History,
} from './history.js'
import { createSheet, setCell, type WorkbookModel } from './model.js'

const free = (): number => 0
const states = (history: History<number>): number[] => history.past.map((entry) => entry.state)

describe('edit history', () => {
  it('undoes and redoes in order, and a new edit discards what was undone', () => {
    const history = recordEdit(recordEdit(emptyHistory<number>(), 1, 2, free), 2, 3, free)
    const undone = undoStep(history, 3, free)!
    expect(undone.state).toBe(2)
    const redone = redoStep(undone.history, 2, free)!
    expect(redone.state).toBe(3)
    expect(states(redone.history)).toEqual([1, 2])
    expect(recordEdit(undone.history, 2, 4, free).future).toEqual([])
  })

  it('with no steps, there is nothing to undo or redo', () => {
    expect(undoStep(emptyHistory<number>(), 1, free)).toBeNull()
    expect(redoStep(emptyHistory<number>(), 1, free)).toBeNull()
  })

  it('keeps at most the step limit', () => {
    let history = emptyHistory<number>()
    for (let step = 0; step < MAX_UNDO_STEPS + 5; step++) history = recordEdit(history, step, step + 1, free)
    expect(history.past).toHaveLength(MAX_UNDO_STEPS)
    expect(history.past[0]?.state).toBe(5)
  })

  it('in a large sheet, the cell limit drops old steps and keeps the last one', () => {
    const heavy = (): number => MAX_UNDO_CELLS / 4
    let history = emptyHistory<number>()
    for (let step = 0; step < 10; step++) history = recordEdit(history, step, step + 1, heavy)
    expect(states(history)).toEqual([6, 7, 8, 9])

    const huge = (): number => MAX_UNDO_CELLS * 2
    expect(states(recordEdit(history, 10, 11, huge))).toEqual([10])
  })

  it('counts only the cells of the sheets the edit changed', () => {
    const big = Array.from({ length: 50 }, (_, row) => row).reduce(
      (sheet, row) => setCell(sheet, row, 0, { value: row }),
      createSheet('Dados'),
    )
    const before: WorkbookModel = { sheets: [big, createSheet('Outra')], activeSheet: 1 }
    const edited = setCell(before.sheets[1]!, 0, 0, { value: 'x' })
    const after: WorkbookModel = { ...before, sheets: [big, edited] }
    expect(changedCells(before, after)).toBe(0)
    expect(changedCells(after, before)).toBe(1)
  })
})
