import { create } from 'zustand'
import { DocumentKind, type DraftSummary } from '@shared/types.js'
import type { IpcResult } from '@shared/ipc.js'
import { canPlaceTab, findOpenTab, isHomeTab, MAX_OPEN_TABS } from '@services/tabs/tab-order.js'
import { t } from '../i18n.js'
import { callReporting, releaseFile } from './context.js'
import { refreshRecents, setRecents } from './recents.js'
import {
  activateTab,
  activeWorkspaceStore,
  addTab,
  removeTab,
  tabOrderOf,
  tabStore,
  useTabs,
} from './tabs.js'
import { createWorkspaceStore, type WorkspaceStore } from './workspace-store.js'
import type { OpenedFile, WorkspaceState } from './types.js'

/** What belongs to the window, not to a tab. */
interface ShellState {
  templateGallery: boolean
  /** The banner offers it until the user decides. */
  pendingDraft: DraftSummary | null
}

export const useShell = create<ShellState>(() => ({ templateGallery: false, pendingDraft: null }))

type NewKind = typeof DocumentKind.Document | typeof DocumentKind.Spreadsheet

/**
 * Which tab a new one may take the place of. A blank new file replaces only Home; a file that
 * already exists also replaces an untouched blank one, as Word does.
 */
type Replaces = 'home' | 'pristine'

const PLATFORM = typeof navigator !== 'undefined' && /Windows/.test(navigator.userAgent) ? 'win32' : 'other'

export function setTemplateGallery(templateGallery: boolean): void {
  useShell.setState({ templateGallery })
}

/** Never saved, never edited, from no template: nothing in it to keep. */
export function isPristine(state: WorkspaceState): boolean {
  const { file } = state
  if (file === null || file.kind === DocumentKind.Pdf) return false
  return file.path === null && file.origin === undefined && !state.isDirty
}

/** On the tab on screen: that is where its banner shows the failure. */
function call<T>(operation: () => Promise<IpcResult<T>>): Promise<T | null> {
  return callReporting(activeWorkspaceStore().setState, operation)
}

function replaceableTab(replaces: Replaces): string | null {
  const { activeId } = useTabs.getState()
  if (isHomeTab(activeId)) return activeId
  if (replaces === 'home') return null
  return isPristine(activeWorkspaceStore().getState()) ? activeId : null
}

/** Before any dialog: picking a file only to hear there is no room would waste the choice. */
function hasRoom(replaces: Replaces): boolean {
  if (canPlaceTab(tabOrderOf(useTabs.getState()), replaceableTab(replaces))) return true
  activeWorkspaceStore()
    .getState()
    .showError({ code: 'INTERNAL', message: t('shell.tabs.full', { count: MAX_OPEN_TABS }) })
  return false
}

function place(store: WorkspaceStore, replaces: Replaces): void {
  const replaced = replaceableTab(replaces)
  const left = replaced === null || isHomeTab(replaced) ? undefined : tabStore(replaced)
  addTab(store, replaced)
  if (left !== undefined) void left.getState().release()
}

/** The same file in two tabs would save over itself: opening it again shows its tab. */
function showOpenTab(path: string): boolean {
  const tabs = useTabs
    .getState()
    .tabs.map((tab) => ({ id: tab.id, path: tab.store.getState().file?.path ?? null }))
  const found = findOpenTab(tabs, path, PLATFORM)
  if (found === null) return false
  activateTab(found)
  return true
}

function reportFailure(store: WorkspaceStore): void {
  const { error } = store.getState()
  if (error !== null) activeWorkspaceStore().getState().showError(error)
}

async function openFile(load: () => Promise<OpenedFile | null>): Promise<boolean> {
  if (!hasRoom('pristine')) return false
  const opened = await load()
  if (opened === null) return false
  if (opened.template !== true && showOpenTab(opened.path)) return true

  const store = createWorkspaceStore()
  if (!store.getState().loadFile(opened)) {
    releaseFile(opened.path)
    reportFailure(store)
    return false
  }
  place(store, 'pristine')
  await refreshRecents()
  return true
}

function startNew(kind: NewKind): void {
  if (!hasRoom('home')) return
  const store = createWorkspaceStore()
  store.getState().startNew(kind)
  place(store, 'home')
}

export function newDocument(): void {
  startNew(DocumentKind.Document)
}

export function newSpreadsheet(): void {
  startNew(DocumentKind.Spreadsheet)
}

export async function openViaDialog(): Promise<boolean> {
  return openFile(async () => {
    const data = await call(() => window.api.file.open({}))
    return data === null || data.canceled ? null : data.file
  })
}

/** Also the files main receives from the file manager. */
export async function openRecent(path: string): Promise<boolean> {
  if (showOpenTab(path)) return true
  return openFile(async () => {
    const data = await call(() => window.api.file.openRecent({ path }))
    if (data === null) {
      await refreshRecents()
      return null
    }
    return data.file
  })
}

/** `null` browses for a `.dotx` or `.ott`. Returns whether the new document reached the screen. */
export async function newFromTemplate(
  template: { readonly source: 'builtin' | 'user'; readonly id: string } | null,
): Promise<boolean> {
  return openFile(async () => {
    if (template !== null) {
      const data = await call(() => window.api.template.open(template))
      return data === null ? null : data.file
    }
    const data = await call(() => window.api.template.browse({}))
    return data === null || data.canceled ? null : data.file
  })
}

export async function clearRecents(): Promise<void> {
  const data = await call(() => window.api.recent.clear({}))
  if (data !== null) setRecents(data.files)
}

export async function checkRecovery(): Promise<void> {
  const result = await window.api.recovery.peek({})
  const latest = result.ok ? result.data.drafts[0] : undefined
  if (latest !== undefined) useShell.setState({ pendingDraft: latest })
}

export async function recoverDraft(): Promise<void> {
  const pending = useShell.getState().pendingDraft
  if (pending === null || !hasRoom('pristine')) return

  useShell.setState({ pendingDraft: null })
  const data = await call(() => window.api.recovery.restore({ draftId: pending.draftId }))
  if (data === null || data.draft === null) return

  const store = createWorkspaceStore()
  if (!store.getState().restoreDraft(data.draft)) {
    reportFailure(store)
    return
  }
  place(store, 'pristine')
}

export async function dismissDraft(): Promise<void> {
  const pending = useShell.getState().pendingDraft
  if (pending === null) return

  useShell.setState({ pendingDraft: null })
  await window.api.recovery.discard({ draftId: pending.draftId })
}

/** `false` when the user kept the tab. */
export async function closeTab(id: string): Promise<boolean> {
  const store = tabStore(id)
  if (store === undefined) return true
  if (isHomeTab(id)) {
    if (useTabs.getState().tabs.length > 1) removeTab(id)
    return true
  }

  // The question names the file; showing its tab first leaves no doubt about which one it is.
  if (store.getState().isDirty) activateTab(id)
  if (!(await store.getState().confirmClose())) return false
  await store.getState().release()
  removeTab(id)
  return true
}

export function closeActiveTab(): Promise<boolean> {
  return closeTab(useTabs.getState().activeId)
}

/** In tab order, each shown before its save; the first that fails or is canceled stops the rest. */
export async function saveAllTabs(): Promise<boolean> {
  for (const tab of useTabs.getState().tabs) {
    if (!tab.store.getState().isDirty) continue
    activateTab(tab.id)
    if (!(await tab.store.getState().save())) return false
  }
  return true
}

export function anyTabDirty(): boolean {
  return useTabs.getState().tabs.some((tab) => tab.store.getState().isDirty)
}
