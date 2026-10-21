import { describe, expect, it } from 'vitest'
import type { LossInventory } from '@shared/types.js'
import { hasReportableLoss, locksEditing, lostOnSave } from './inventory.js'

const inventory = (partial: Partial<LossInventory>): LossInventory => ({
  lost: [],
  invisible: [],
  structural: [],
  ...partial,
})

describe('something to warn about', () => {
  it('no, when the file opened whole', () => {
    expect(hasReportableLoss(inventory({}))).toBe(false)
    expect(hasReportableLoss(undefined)).toBe(false)
  })

  it('yes, both for what is lost and for what is not shown', () => {
    expect(hasReportableLoss(inventory({ lost: ['comentários'] }))).toBe(true)
    expect(hasReportableLoss(inventory({ invisible: ['notas de rodapé'] }))).toBe(true)
  })
})

describe('read-only', () => {
  it('locks for what is lost when editing the block that anchors it', () => {
    expect(locksEditing(inventory({ structural: ['comentários'] }))).toBe(true)
  })

  it('does not lock for a loss of appearance', () => {
    // Almost every corpus document has an anchored image: locking for that would lock everyday use,
    // and the user would learn to unlock without reading.
    expect(locksEditing(inventory({ lost: ['posicionamento de imagem'] }))).toBe(false)
    expect(locksEditing(undefined)).toBe(false)
  })
})

describe('what a save lost', () => {
  it('is only what was lost: the invisible stays in the file', () => {
    expect(lostOnSave(inventory({ lost: ['mesclagem vertical'], invisible: ['formas'] }))).toEqual([
      'mesclagem vertical',
    ])
  })

  it('reports each loss once, even if it happened in several cells', () => {
    expect(lostOnSave(inventory({ lost: ['mesclagem vertical', 'mesclagem vertical'] }))).toEqual([
      'mesclagem vertical',
    ])
  })

  it('is nothing when the format has no inventory', () => {
    expect(lostOnSave(undefined)).toEqual([])
  })
})
