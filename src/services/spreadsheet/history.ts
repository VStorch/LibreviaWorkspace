import type { WorkbookModel } from './model.js'

/**
 * Each step keeps a whole previous state: the model is immutable, so what it costs is what changed,
 * which `Weigh` measures. The bound is on that cost as well as on the number of steps.
 */
export interface Entry<State> {
  readonly state: State
  readonly cost: number
}

export interface History<State> {
  readonly past: readonly Entry<State>[]
  readonly future: readonly Entry<State>[]
}

/** What keeping `kept` costs beyond `neighbor`, the state it sits next to. */
export type Weigh<State> = (kept: State, neighbor: State) => number

export const MAX_UNDO_STEPS = 100

/** Cells held by the kept copies: each edit copies the cell map of the sheet it changed. */
export const MAX_UNDO_CELLS = 1_000_000

export function emptyHistory<State>(): History<State> {
  return { past: [], future: [] }
}

/** Sheets shared with the neighbor cost nothing: only a changed sheet's map was copied. */
export function changedCells(kept: WorkbookModel, neighbor: WorkbookModel): number {
  const shared = new Set(neighbor.sheets)
  return kept.sheets
    .filter((sheet) => !shared.has(sheet))
    .reduce((total, sheet) => total + Object.keys(sheet.cells).length, 0)
}

const costOf = <State>(entries: readonly Entry<State>[]): number =>
  entries.reduce((total, entry) => total + entry.cost, 0)

function isOverLimit<State>(past: readonly Entry<State>[], future: readonly Entry<State>[]): boolean {
  return past.length + future.length > MAX_UNDO_STEPS || costOf(past) + costOf(future) > MAX_UNDO_CELLS
}

/** The oldest steps go first; the nearest step each way stays, even over the limit. */
function bounded<State>(history: History<State>): History<State> {
  let { past, future } = history
  while (past.length > 1 && isOverLimit(past, future)) past = past.slice(1)
  while (future.length > 1 && isOverLimit(past, future)) future = future.slice(0, -1)
  return { past, future }
}

/** A new edit drops what had been undone, as in any editor. */
export function recordEdit<State>(
  history: History<State>,
  previous: State,
  next: State,
  weigh: Weigh<State>,
): History<State> {
  const entry = { state: previous, cost: weigh(previous, next) }
  return bounded({ past: [...history.past, entry], future: [] })
}

export interface Step<State> {
  readonly history: History<State>
  readonly state: State
}

export function undoStep<State>(
  history: History<State>,
  current: State,
  weigh: Weigh<State>,
): Step<State> | null {
  const entry = history.past.at(-1)
  if (entry === undefined) return null
  const kept = { state: current, cost: weigh(current, entry.state) }
  return {
    state: entry.state,
    history: bounded({ past: history.past.slice(0, -1), future: [kept, ...history.future] }),
  }
}

export function redoStep<State>(
  history: History<State>,
  current: State,
  weigh: Weigh<State>,
): Step<State> | null {
  const entry = history.future[0]
  if (entry === undefined) return null
  const kept = { state: current, cost: weigh(current, entry.state) }
  return {
    state: entry.state,
    history: bounded({ past: [...history.past, kept], future: history.future.slice(1) }),
  }
}
