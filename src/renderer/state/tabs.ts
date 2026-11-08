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
