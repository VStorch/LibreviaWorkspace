import type { LossInventory } from '@shared/types.js'

/** Um alerta que abre em todo arquivo é fechado sem ler. */
export function hasReportableLoss(inventory: LossInventory | undefined): boolean {
  return inventory !== undefined && (inventory.invisible.length > 0 || inventory.lost.length > 0)
}

/** Perda de aparência não trava: quase todo documento do corpus tem alguma. */
export function locksEditing(inventory: LossInventory | undefined): boolean {
  return inventory !== undefined && inventory.structural.length > 0
}

/** Só o `lost`, cada coisa uma vez: o sidecar registra por célula e por bloco. */
export function lostOnSave(inventory: LossInventory | undefined): readonly string[] {
  return inventory === undefined ? [] : [...new Set(inventory.lost)]
}
