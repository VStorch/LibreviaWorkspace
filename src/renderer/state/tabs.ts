import { useState } from 'react'
import { create } from 'zustand'
import { createWorkspaceStore, type WorkspaceStore } from './workspace-store.js'

export interface TabEntry {
  id: string
  store: WorkspaceStore
}

export interface TabsState {
  tabs: TabEntry[]
  activeId: string
}

const FIRST_TAB_ID = 'tab-1'

export const useTabs = create<TabsState>(() => ({
  tabs: [{ id: FIRST_TAB_ID, store: createWorkspaceStore() }],
  activeId: FIRST_TAB_ID,
}))

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
