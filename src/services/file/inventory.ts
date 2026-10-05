import type { LossInventory } from '@shared/types.js'

/** A warning that opens on every file is closed without reading. */
export function hasReportableLoss(inventory: LossInventory | undefined): boolean {
  return inventory !== undefined && (inventory.invisible.length > 0 || inventory.lost.length > 0)
}

/** Appearance loss does not lock: almost every corpus document has some. */
export function locksEditing(inventory: LossInventory | undefined): boolean {
  return inventory !== undefined && inventory.structural.length > 0
}

/** Only `lost`, each thing once: the sidecar records per cell and per block. */
export function lostOnSave(inventory: LossInventory | undefined): readonly string[] {
  return inventory === undefined ? [] : [...new Set(inventory.lost)]
}
