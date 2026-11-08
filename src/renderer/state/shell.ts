import { create } from 'zustand'
import { DocumentKind, type DraftSummary } from '@shared/types.js'
import type { IpcResult } from '@shared/ipc.js'
import { canPlaceTab, isHomeTab, MAX_OPEN_TABS } from '@services/tabs/tab-order.js'
import { saveAllPlan, type DirtyTab } from '@services/tabs/quit.js'
import { t } from '../i18n.js'
import { callReporting, releaseFile } from './context.js'
import { refreshRecents, setRecents } from './recents.js'
import {
  activateTab,
  activeWorkspaceStore,
  addTab,
  findTabWithPath,
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
  const found = findTabWithPath(path)
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
    releaseFile(opened.path, store.getState())
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

function openFilePaths(): string[] {
  return useTabs.getState().tabs.flatMap((tab) => tab.store.getState().file?.path ?? [])
}

/** Main does not read again a file already open: that tab's package would change under it. */
export async function openViaDialog(): Promise<boolean> {
  return openFile(async () => {
    const data = await call(() => window.api.file.open({ openPaths: openFilePaths() }))
    if (data === null || data.canceled) return null
    if ('alreadyOpen' in data) {
      showOpenTab(data.alreadyOpen)
      return null
    }
    return data.file
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

/** The drafts of the tabs open now are theirs, not left by a crash. */
export async function checkRecovery(): Promise<void> {
  const result = await window.api.recovery.peek({})
  if (!result.ok) return
  const live = new Set(useTabs.getState().tabs.map((tab) => tab.store.getState().draftId))
  const latest = result.data.drafts.find((draft) => !live.has(draft.draftId))
  useShell.setState({ pendingDraft: latest ?? null })
}

/** The file in a tab and its draft in another would each save over the other. */
function refuseOpenFile(pending: DraftSummary): boolean {
  if (pending.path === null || !showOpenTab(pending.path)) return false
  activeWorkspaceStore()
    .getState()
    .showError({ code: 'INTERNAL', message: t('shell.recovery.fileOpen', { name: pending.name }) })
  return true
}

export async function recoverDraft(): Promise<void> {
  const pending = useShell.getState().pendingDraft
  if (pending === null || refuseOpenFile(pending) || !hasRoom('pristine')) return

  useShell.setState({ pendingDraft: null })
  const data = await call(() => window.api.recovery.restore({ draftId: pending.draftId }))
  if (data !== null && data.draft !== null) {
    const store = createWorkspaceStore()
    if (store.getState().restoreDraft(data.draft)) place(store, 'pristine')
    else reportFailure(store)
  }
  await checkRecovery()
}

export async function dismissDraft(): Promise<void> {
  const pending = useShell.getState().pendingDraft
  if (pending === null) return

  useShell.setState({ pendingDraft: null })
  await window.api.recovery.discard({ draftId: pending.draftId })
  await checkRecovery()
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

/** In tab order; the first tab the user keeps stops the rest. */
export async function closeAllTabs(): Promise<boolean> {
  for (const { id } of useTabs.getState().tabs) {
    if (!(await closeTab(id))) return false
  }
  return true
}

/**
 * In tab order; the first that fails or is canceled stops the rest, and its tab comes to the
 * screen, where its banner says why.
 */
export async function saveAllTabs(): Promise<boolean> {
  const candidates = useTabs.getState().tabs.map(({ id, store }) => {
    const { isDirty, file } = store.getState()
    return { id, dirty: isDirty, hasPath: (file?.path ?? null) !== null }
  })
  for (const step of saveAllPlan(candidates)) {
    const store = tabStore(step.id)
    if (store === undefined) continue
    if (step.showFirst) activateTab(step.id)
    if (await store.getState().save()) continue
    activateTab(step.id)
    return false
  }
  return true
}

/** Every modified tab, in tab order, by the name its tab shows. */
export function dirtyTabs(): DirtyTab[] {
  return useTabs.getState().tabs.flatMap(({ id, store }) => {
    const { isDirty, file } = store.getState()
    return isDirty ? [{ id, label: file?.name ?? t('shell.file.untitled') }] : []
  })
}
