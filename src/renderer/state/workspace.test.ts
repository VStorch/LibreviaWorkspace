import { afterEach, describe, expect, it, vi } from 'vitest'
import { createElement } from 'react'
import { renderToString } from 'react-dom/server'
import { useTabs } from './tabs.js'
import { createWorkspaceStore, useWorkspace, type WorkspaceStore } from './workspace.js'

const initialTabs = useTabs.getState()

function openSecondTab(): WorkspaceStore {
  const store = createWorkspaceStore()
  useTabs.setState({ tabs: [...initialTabs.tabs, { id: 'tab-2', store }] })
  return store
}

function activate(id: string): void {
  useTabs.setState({ activeId: id })
}

function selectedDuringRender<T>(selector: Parameters<typeof useWorkspace<T>>[0]): T {
  let selected: T | undefined
  function Probe(): null {
    selected = useWorkspace(selector)
    return null
  }
  renderToString(createElement(Probe))
  return selected as T
}

afterEach(() => {
  vi.restoreAllMocks()
  useTabs.setState(initialTabs, true)
})

describe('useWorkspace', () => {
  it('reads the selector from the active tab store', () => {
    openSecondTab()
    const first = initialTabs.tabs[0]!.store
    // Server rendering reads each store's initial state, so the active store's is the probe.
    vi.spyOn(first, 'getInitialState').mockReturnValue({ ...first.getState(), pageCount: 9 })

    expect(selectedDuringRender((state) => state.pageCount)).toBe(9)
  })

  it('reads and writes the active store through getState and setState', () => {
    const second = openSecondTab()
    activate('tab-2')

    useWorkspace.setState({ readOnly: true })

    expect(useWorkspace.getState()).toBe(second.getState())
    expect(second.getState().readOnly).toBe(true)
    expect(initialTabs.tabs[0]!.store.getState().readOnly).toBe(false)
  })

  it('moves a subscription to the newly active store', () => {
    const second = openSecondTab()
    const first = initialTabs.tabs[0]!.store
    const listener = vi.fn()
    const unsubscribe = useWorkspace.subscribe(listener)

    activate('tab-2')
    expect(listener).toHaveBeenLastCalledWith(second.getState(), first.getState())
    listener.mockClear()

    first.setState({ busy: true })
    expect(listener).not.toHaveBeenCalled()
    second.setState({ busy: true })
    expect(listener).toHaveBeenCalledOnce()

    unsubscribe()
    second.setState({ busy: false })
    expect(listener).toHaveBeenCalledOnce()
  })

  it('keeps each tab state apart', () => {
    const second = openSecondTab()

    useWorkspace.setState({ pageCount: 7 })

    expect(useWorkspace.getState().pageCount).toBe(7)
    expect(second.getState().pageCount).toBe(1)
    useWorkspace.setState({ pageCount: 1 })
  })
})
