import { describe, expect, it } from 'vitest'
import type { LossInventory } from '@shared/types.js'
import { hasReportableLoss, locksEditing } from './inventory.js'

const inventory = (partial: Partial<LossInventory>): LossInventory => ({
  lost: [],
  invisible: [],
  structural: [],
  ...partial,
})

describe('há o que avisar', () => {
  it('não, quando o arquivo abriu inteiro', () => {
    expect(hasReportableLoss(inventory({}))).toBe(false)
    expect(hasReportableLoss(undefined)).toBe(false)
  })

  it('sim, tanto para o que some quanto para o que não aparece', () => {
    expect(hasReportableLoss(inventory({ lost: ['comentários'] }))).toBe(true)
    expect(hasReportableLoss(inventory({ invisible: ['notas de rodapé'] }))).toBe(true)
  })
})

describe('somente leitura', () => {
  it('trava pelo que some ao editar o bloco que o ancora', () => {
    expect(locksEditing(inventory({ structural: ['comentários'] }))).toBe(true)
  })

  it('não trava por perda de aparência', () => {
    // Quase todo documento do corpus tem uma imagem ancorada: travar por isso
    // travaria o uso do dia a dia, e o usuário aprenderia a liberar sem ler.
    expect(locksEditing(inventory({ lost: ['posicionamento de imagem'] }))).toBe(false)
    expect(locksEditing(undefined)).toBe(false)
  })
})
