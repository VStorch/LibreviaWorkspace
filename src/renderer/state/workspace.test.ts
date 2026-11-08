import { afterEach, describe, expect, it, vi } from 'vitest'
import { createElement } from 'react'
import { renderToString } from 'react-dom/server'
import { DocumentKind } from '@shared/types.js'
import { autosaveAllTabs, useTabs } from './tabs.js'
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
  vi.unstubAllGlobals()
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

describe('work started in a tab', () => {
  it('lands in the tab that started it after a switch', async () => {
    let finish: (value: unknown) => void = () => undefined
    const inspect = vi.fn(() => new Promise((resolve) => (finish = resolve)))
    vi.stubGlobal('window', { api: { privacy: { inspect } } })
    const first = initialTabs.tabs[0]!.store
    first.setState({ file: { path: '/docs/a.docx', name: 'a.docx', kind: DocumentKind.Document } })
    const second = openSecondTab()

    const pending = useWorkspace.getState().inspectPersonalData()
    activate('tab-2')
    finish({ ok: false, error: { code: 'INTERNAL', message: 'failed' } })
    await pending

    expect(first.getState().error?.message).toBe('failed')
    expect(first.getState().busy).toBe(false)
    expect(second.getState().error).toBeNull()
    expect(second.getState().busy).toBe(false)
  })

  it('autosaves every changed tab under its own draft id', async () => {
    const autosave = vi.fn(async (_request: { draftId: string }) => ({ ok: true, data: { savedAt: 1 } }))
    vi.stubGlobal('window', { api: { file: { autosave } } })
    const first = createWorkspaceStore()
    const second = createWorkspaceStore()
    const clean = createWorkspaceStore()
    useTabs.setState({
      tabs: [
        { id: 'tab-a', store: first },
        { id: 'tab-b', store: second },
        { id: 'tab-c', store: clean },
      ],
      activeId: 'tab-a',
    })
    for (const store of [first, second]) {
      store.setState({ file: { path: null, name: 'a.sdoc', kind: DocumentKind.Document }, isDirty: true })
    }

    await autosaveAllTabs()

    const ids = autosave.mock.calls.map(([request]) => request.draftId)
    expect(ids).toEqual([first.getState().draftId, second.getState().draftId])
  })
})
