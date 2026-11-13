import { create } from 'zustand'
import { DocumentKind, MenuCommand, type DraftSummary } from '@shared/types.js'
import type { IpcResult } from '@shared/ipc.js'
import { isOdtTemplatePath, isWordTemplatePath } from '@services/file/formats.js'
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
  showHome,
  tabOrderOf,
  tabStore,
  useTabs,
} from './tabs.js'
import { createWorkspaceStore, type WorkspaceStore } from './workspace-store.js'
import { openInAnotherWindow, opensWindows } from './open-mode.js'
import type { OpenedFile, WorkspaceState } from './types.js'

/** What belongs to the window, not to a tab. */
interface ShellState {
  templateGallery: boolean
  /** The banner offers it until the user decides. */
  pendingDraft: DraftSummary | null
  /** Left by the same crash, offered after `pendingDraft` or all at once. */
  otherDrafts: number
}

export const useShell = create<ShellState>(() => ({
  templateGallery: false,
  pendingDraft: null,
  otherDrafts: 0,
}))

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

const NEW_COMMANDS = {
  [DocumentKind.Document]: MenuCommand.NewDocument,
  [DocumentKind.Spreadsheet]: MenuCommand.NewSpreadsheet,
} as const

/** With one file per window and a file on screen, the new one goes to a window of its own. */
function goesToNewWindow(replaces: Replaces): boolean {
  return opensWindows() && replaceableTab(replaces) === null
}

function startNew(kind: NewKind): void {
  if (goesToNewWindow('home')) {
    void openInAnotherWindow({ kind: 'new', command: NEW_COMMANDS[kind] })
    return
  }
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

/** One by one, in the order picked; a full tab bar stops the rest. */
export async function openViaDialog(): Promise<boolean> {
  const data = await call(() => window.api.file.open({}))
  if (data === null || data.canceled) return false
  let opened = false
  for (const path of data.paths) {
    if (!opensWindows() && !hasRoom('pristine')) break
    opened = (await openPicked(path)) || opened
  }
  return opened
}

/** A template opens a new document, even when a tab edits the template itself. */
async function openPicked(path: string): Promise<boolean> {
  if (!isWordTemplatePath(path) && !isOdtTemplatePath(path)) return openRecent(path)
  if (goesToNewWindow('pristine')) return openInAnotherWindow({ kind: 'path', path, hasRoom: false })
  return openFile(() => readPicked(path))
}

async function readPicked(path: string): Promise<OpenedFile | null> {
  const data = await call(() => window.api.file.openRecent({ path }))
  return data === null ? null : data.file
}

/**
 * Also the files main receives from the file manager or the open dialog. A file already open is not
 * read again: that tab's package would change under it.
 */
export async function openRecent(path: string): Promise<boolean> {
  if (showOpenTab(path)) return true
  if (opensWindows()) {
    const hasRoom = replaceableTab('pristine') !== null
    if (await openInAnotherWindow({ kind: 'path', path, hasRoom })) return true
  }
  return openFile(async () => {
    const opened = await readPicked(path)
    if (opened === null) await refreshRecents()
    return opened
  })
}

/** `null` browses for a `.dotx` or `.ott`. Returns whether the new document reached the screen. */
export async function newFromTemplate(
  template: { readonly source: 'builtin' | 'user'; readonly id: string } | null,
): Promise<boolean> {
  if (goesToNewWindow('pristine')) return openInAnotherWindow({ kind: 'template', template })
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
async function offeredDrafts(): Promise<DraftSummary[] | null> {
  const result = await window.api.recovery.peek({})
  if (!result.ok) return null
  const live = new Set(useTabs.getState().tabs.map((tab) => tab.store.getState().draftId))
  return result.data.drafts.filter((draft) => !live.has(draft.draftId))
}

export async function checkRecovery(): Promise<void> {
  const drafts = await offeredDrafts()
  if (drafts === null) return
  useShell.setState({ pendingDraft: drafts[0] ?? null, otherDrafts: Math.max(drafts.length - 1, 0) })
}

/** The file in a tab and its draft in another would each save over the other. */
function refuseOpenFile(pending: DraftSummary): boolean {
  if (pending.path === null || !showOpenTab(pending.path)) return false
  activeWorkspaceStore()
    .getState()
    .showError({ code: 'INTERNAL', message: t('shell.recovery.fileOpen', { name: pending.name }) })
  return true
}

type Restored = 'restored' | 'refused' | 'full'

async function restoreDraft(pending: DraftSummary): Promise<Restored> {
  if (refuseOpenFile(pending)) return 'refused'
  if (!hasRoom('pristine')) return 'full'
  const data = await call(() => window.api.recovery.restore({ draftId: pending.draftId }))
  if (data !== null && data.draft !== null) {
    const store = createWorkspaceStore()
    if (store.getState().restoreDraft(data.draft)) place(store, 'pristine')
    else reportFailure(store)
  }
  return 'restored'
}

export async function recoverDraft(): Promise<void> {
  const pending = useShell.getState().pendingDraft
  if (pending === null) return
  // Hidden while it loads, so a second click does not restore it twice.
  useShell.setState({ pendingDraft: null })
  if ((await restoreDraft(pending)) === 'restored') await checkRecovery()
  else useShell.setState({ pendingDraft: pending })
}

/** Newest first; a draft whose file is open stays on the banner, and a full tab bar stops the rest. */
export async function recoverAllDrafts(): Promise<void> {
  const drafts = await offeredDrafts()
  if (drafts === null) return
  useShell.setState({ pendingDraft: null, otherDrafts: 0 })
  for (const draft of drafts) {
    if ((await restoreDraft(draft)) === 'full') break
  }
  await checkRecovery()
}

export async function dismissDraft(): Promise<void> {
  const pending = useShell.getState().pendingDraft
  if (pending === null) return

  useShell.setState({ pendingDraft: null, otherDrafts: 0 })
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

/** Ctrl+T and the tab bar's +: with one file per window and a file on screen, Home opens a window. */
export function openHome(): void {
  if (goesToNewWindow('home')) void openInAnotherWindow({ kind: 'new', command: MenuCommand.ShowHome })
  else showHome()
}
