import type { WorkspaceState } from './types.js'

/** Whether a tab other than the asker's has the path open. */
type PathHolder = (path: string, asker: WorkspaceState) => boolean

let holder: PathHolder = () => false

/** Set by the tab list: the stores cannot import it, since it creates them. */
export function setPathHolder(next: PathHolder): void {
  holder = next
}

export function isOpenInAnotherTab(path: string, asker: WorkspaceState): boolean {
  return holder(path, asker)
}
