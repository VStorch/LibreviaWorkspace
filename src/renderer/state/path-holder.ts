import { comparablePath } from '@services/tabs/tab-order.js'
import type { OpenFile, WorkspaceState } from './types.js'

/** The files of the tabs other than the asker's. */
type OtherFiles = (asker: WorkspaceState) => readonly OpenFile[]

let otherFiles: OtherFiles = () => []

/** Set by the tab list: the stores cannot import it, since it creates them. */
export function setOtherFiles(next: OtherFiles): void {
  otherFiles = next
}

export const PLATFORM =
  typeof navigator !== 'undefined' && /Windows/.test(navigator.userAgent) ? 'win32' : 'other'

/** On the platform's terms, as main compares the paths it keeps. */
export function isSamePath(a: string, b: string | null | undefined): boolean {
  return typeof b === 'string' && comparablePath(a, PLATFORM) === comparablePath(b, PLATFORM)
}

export function isOpenInAnotherTab(path: string, asker: WorkspaceState): boolean {
  return otherFiles(asker).some((file) => isSamePath(path, file.path))
}

/** A document made from a template saves from the template's package, kept under its path. */
export function isNeededByAnotherTab(path: string, asker: WorkspaceState): boolean {
  return otherFiles(asker).some((file) => isSamePath(path, file.path) || isSamePath(path, file.origin))
}
