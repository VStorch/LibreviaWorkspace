import { useState } from 'react'
import { create } from 'zustand'
import {
  activateTab as activateInOrder,
  closeTab as closeInOrder,
  findOpenTab,
  HOME_TAB_ID,
  nextTab,
  openTab,
  previousTab,
  placeTab,
  type TabOrder,
} from '@services/tabs/tab-order.js'
import { setPathHolder } from './path-holder.js'
import { createWorkspaceStore, type WorkspaceStore } from './workspace-store.js'

export interface TabEntry {
  id: string
  store: WorkspaceStore
}

export interface TabsState {
  tabs: TabEntry[]
  activeId: string
}

/** Home is a tab like the others, with a store that never holds a file. */
const HOME: TabEntry = { id: HOME_TAB_ID, store: createWorkspaceStore() }

export const useTabs = create<TabsState>(() => ({ tabs: [HOME], activeId: HOME_TAB_ID }))

let lastTabNumber = 0

const PLATFORM = typeof navigator !== 'undefined' && /Windows/.test(navigator.userAgent) ? 'win32' : 'other'

function tabPaths(tabs: readonly TabEntry[]): { id: string; path: string | null }[] {
  return tabs.map((tab) => ({ id: tab.id, path: tab.store.getState().file?.path ?? null }))
}

/** The tab with the file open, on the platform's terms of what makes two paths the same. */
export function findTabWithPath(path: string): string | null {
  return findOpenTab(tabPaths(useTabs.getState().tabs), path, PLATFORM)
}

setPathHolder((path, asker) => {
  const others = useTabs.getState().tabs.filter((tab) => tab.store.getState() !== asker)
  return findOpenTab(tabPaths(others), path, PLATFORM) !== null
})

export function tabOrderOf({ tabs, activeId }: TabsState): TabOrder {
  return { ids: tabs.map((tab) => tab.id), activeId }
}

/** The tab being left keeps what its editor shows: hidden, it has no editor to read from. */
function applyOrder(order: TabOrder, added?: TabEntry): void {
  const current = useTabs.getState()
  if (order.activeId !== current.activeId) activeStoreOf(current).getState().suspend()
  const known = added === undefined ? current.tabs : [...current.tabs, added]
  const tabs = order.ids.map((id) => known.find((tab) => tab.id === id) ?? HOME)
  useTabs.setState({ tabs, activeId: order.activeId })
}

export function activateTab(id: string): void {
  applyOrder(activateInOrder(tabOrderOf(useTabs.getState()), id))
}

export function activateNextTab(): void {
  applyOrder(nextTab(tabOrderOf(useTabs.getState())))
}

export function activatePreviousTab(): void {
  applyOrder(previousTab(tabOrderOf(useTabs.getState())))
}

export function showHome(): void {
  HOME.store.setState({ error: null })
  applyOrder(openTab(tabOrderOf(useTabs.getState()), HOME_TAB_ID))
}

/** Activates the new tab, in the place of `replacedId` when given. */
export function addTab(store: WorkspaceStore, replacedId: string | null): void {
  lastTabNumber += 1
  const entry = { id: `tab-${lastTabNumber}`, store }
  applyOrder(placeTab(tabOrderOf(useTabs.getState()), entry.id, replacedId), entry)
}

/** Without asking: the caller already dealt with the changes. Closing the last leaves Home alone. */
export function removeTab(id: string): void {
  applyOrder(closeInOrder(tabOrderOf(useTabs.getState()), id))
}

export function tabStore(id: string): WorkspaceStore | undefined {
  return useTabs.getState().tabs.find((tab) => tab.id === id)?.store
}

export function activeStoreOf({ tabs, activeId }: TabsState): WorkspaceStore {
  const active = tabs.find((tab) => tab.id === activeId) ?? tabs[0]
  if (active === undefined) throw new Error('No tab is open')
  return active.store
}

export function activeWorkspaceStore(): WorkspaceStore {
  return activeStoreOf(useTabs.getState())
}

/** Every tab, hidden ones included: a hidden tab reads its content from its own snapshot. */
export async function autosaveAllTabs(): Promise<void> {
  await Promise.all(useTabs.getState().tabs.map((tab) => tab.store.getState().autosave()))
}

/** The tab a mounted view belongs to: the one active when it mounted, since `App` remounts per tab. */
export function useOwnTabId(): string {
  const [id] = useState(() => useTabs.getState().activeId)
  return id
}

/** A value each tab's view holds while mounted; readers get the active tab's. */
export interface TabSlot<T> {
  /** Returns the release function, which leaves a newer value in place. */
  hold(tabId: string, value: T): () => void
  active(): T | undefined
}

export function createTabSlot<T>(): TabSlot<T> {
  const values = new Map<string, T>()
  return {
    hold(tabId, value) {
      values.set(tabId, value)
      return () => {
        if (values.get(tabId) === value) values.delete(tabId)
      }
    },
    active: () => values.get(useTabs.getState().activeId),
  }
}
