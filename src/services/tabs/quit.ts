export interface DirtyTab {
  readonly id: string
  readonly label: string
}

export type QuitQuestion =
  | { readonly kind: 'none' }
  | { readonly kind: 'one'; readonly label: string }
  | { readonly kind: 'several'; readonly labels: readonly string[] }

/** One modified file gets the usual question; several get one box with all their names. */
export function quitQuestion(dirty: readonly DirtyTab[]): QuitQuestion {
  const [first] = dirty
  if (first === undefined) return { kind: 'none' }
  if (dirty.length === 1) return { kind: 'one', label: first.label }
  return { kind: 'several', labels: dirty.map((tab) => tab.label) }
}

export interface SaveCandidate {
  readonly id: string
  readonly dirty: boolean
  readonly hasPath: boolean
}

export interface SaveStep {
  readonly id: string
  /** A file never saved asks where to go: its tab is shown first, so the dialog has a face. */
  readonly showFirst: boolean
}

/** In tab order, only the modified ones; the caller stops at the first that fails or is canceled. */
export function saveAllPlan(tabs: readonly SaveCandidate[]): SaveStep[] {
  return tabs.filter((tab) => tab.dirty).map((tab) => ({ id: tab.id, showFirst: !tab.hasPath }))
}
