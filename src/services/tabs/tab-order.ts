export const HOME_TAB_ID = 'home'
export const MAX_OPEN_TABS = 16

export interface TabOrder {
  readonly ids: readonly string[]
  readonly activeId: string
}

export interface TabPath {
  readonly id: string
  readonly path: string | null
}

export const HOME_ONLY: TabOrder = { ids: [HOME_TAB_ID], activeId: HOME_TAB_ID }

export function isHomeTab(id: string): boolean {
  return id === HOME_TAB_ID
}

export function openTabCount(order: TabOrder): number {
  return order.ids.filter((id) => !isHomeTab(id)).length
}

export function canOpenAnother(order: TabOrder): boolean {
  return openTabCount(order) < MAX_OPEN_TABS
}

export function openTab(order: TabOrder, id: string): TabOrder {
  if (order.ids.includes(id)) return activateTab(order, id)
  return { ids: [...order.ids, id], activeId: id }
}

export function activateTab(order: TabOrder, id: string): TabOrder {
  return order.ids.includes(id) ? { ...order, activeId: id } : order
}

export function closeTab(order: TabOrder, id: string): TabOrder {
  const index = order.ids.indexOf(id)
  if (index < 0) return order
  const ids = order.ids.filter((other) => other !== id)
  if (ids.length === 0) return HOME_ONLY
  if (order.activeId !== id) return { ids, activeId: order.activeId }
  return { ids, activeId: ids[Math.min(index, ids.length - 1)] ?? HOME_TAB_ID }
}

export function nextTab(order: TabOrder): TabOrder {
  return stepTab(order, 1)
}

export function previousTab(order: TabOrder): TabOrder {
  return stepTab(order, -1)
}

function stepTab(order: TabOrder, step: 1 | -1): TabOrder {
  const count = order.ids.length
  const index = order.ids.indexOf(order.activeId)
  if (count === 0 || index < 0) return order
  return { ...order, activeId: order.ids[(index + step + count) % count] ?? order.activeId }
}

export function moveTab(order: TabOrder, id: string, toIndex: number): TabOrder {
  const from = order.ids.indexOf(id)
  if (from < 0) return order
  const ids = order.ids.filter((other) => other !== id)
  const target = Math.max(0, Math.min(Math.trunc(toIndex), ids.length))
  ids.splice(target, 0, id)
  return { ...order, ids }
}

export function findOpenTab(tabs: readonly TabPath[], path: string, platform: string): string | null {
  const wanted = comparablePath(path, platform)
  const found = tabs.find((tab) => tab.path !== null && comparablePath(tab.path, platform) === wanted)
  return found?.id ?? null
}

export function comparablePath(path: string, platform: string): string {
  const windows = platform === 'win32'
  const slashed = windows ? path.replace(/\\/g, '/') : path
  const uncPrefix = windows && slashed.startsWith('//') ? '/' : ''
  const collapsed = uncPrefix + slashed.replace(/\/{2,}/g, '/')
  const trimmed = collapsed.length > 1 ? collapsed.replace(/\/+$/, '') : collapsed
  const rooted = windows && /^[A-Za-z]:$/.test(trimmed) ? `${trimmed}/` : trimmed
  return windows ? rooted.toLowerCase() : rooted
}
