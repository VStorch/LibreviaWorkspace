import { describe, expect, it } from 'vitest'
import { quitQuestion, saveAllPlan } from './quit.js'

describe('the question on quitting', () => {
  it('asks nothing when no tab is modified', () => {
    expect(quitQuestion([])).toEqual({ kind: 'none' })
  })

  it('asks about the one file by name', () => {
    expect(quitQuestion([{ id: 'tab-1', label: 'ata.docx' }])).toEqual({ kind: 'one', label: 'ata.docx' })
  })

  it('lists every modified file, in tab order, in a single question', () => {
    const question = quitQuestion([
      { id: 'tab-2', label: 'b.docx' },
      { id: 'tab-1', label: 'a.xlsx' },
    ])

    expect(question).toEqual({ kind: 'several', labels: ['b.docx', 'a.xlsx'] })
  })
})

describe('saving every tab', () => {
  it('saves only the modified tabs, in tab order', () => {
    const plan = saveAllPlan([
      { id: 'tab-1', dirty: true, hasPath: true },
      { id: 'tab-2', dirty: false, hasPath: true },
      { id: 'tab-3', dirty: true, hasPath: true },
    ])

    expect(plan.map((step) => step.id)).toEqual(['tab-1', 'tab-3'])
  })

  it('shows a tab before its save dialog only when the file was never saved', () => {
    const plan = saveAllPlan([
      { id: 'tab-1', dirty: true, hasPath: true },
      { id: 'tab-2', dirty: true, hasPath: false },
    ])

    expect(plan).toEqual([
      { id: 'tab-1', showFirst: false },
      { id: 'tab-2', showFirst: true },
    ])
  })
})
