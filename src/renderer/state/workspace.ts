import { useStore } from 'zustand'
import { activeStoreOf, activeWorkspaceStore, useTabs } from './tabs.js'
import type { WorkspaceStore } from './workspace-store.js'
import type { WorkspaceState } from './types.js'

export type { DocumentSource, WorkspaceState } from './types.js'
export { createWorkspaceStore, type WorkspaceStore } from './workspace-store.js'

type Listener = Parameters<WorkspaceStore['subscribe']>[0]

interface UseWorkspace extends WorkspaceStore {
  (): WorkspaceState
  <T>(selector: (state: WorkspaceState) => T): T
}

function useActiveWorkspace<T>(selector?: (state: WorkspaceState) => T): T | WorkspaceState {
  const store = useTabs(activeStoreOf)
  return useStore<WorkspaceStore, T | WorkspaceState>(store, selector ?? ((state) => state))
}

/** Follows the active tab: the listener moves with it and hears the switch as a change. */
function subscribeToActive(listener: Listener): () => void {
  let current = activeWorkspaceStore()
  let detach = current.subscribe(listener)
  const detachTabs = useTabs.subscribe((tabs) => {
    const next = activeStoreOf(tabs)
    if (next === current) return
    const previous = current.getState()
    detach()
    current = next
    detach = next.subscribe(listener)
    listener(next.getState(), previous)
  })
  return () => {
    detachTabs()
    detach()
  }
}

/** The active tab's workspace, with the same shape the single store had. */
export const useWorkspace = Object.assign(useActiveWorkspace, {
  getState: () => activeWorkspaceStore().getState(),
  getInitialState: () => activeWorkspaceStore().getInitialState(),
  setState: ((partial, replace) =>
    activeWorkspaceStore().setState(
      partial as WorkspaceState,
      replace as true,
    )) as WorkspaceStore['setState'],
  subscribe: subscribeToActive,
}) as UseWorkspace
