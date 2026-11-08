import { afterEach, describe, expect, it, vi } from 'vitest'
import { DocumentKind } from '@shared/types.js'
import type { DocumentNode } from '@services/document/model.js'
import { activateTab, addTab, autosaveAllTabs, createTabSlot, useTabs } from './tabs.js'
import { createWorkspaceStore } from './workspace-store.js'

const initialTabs = useTabs.getState()

afterEach(() => {
  vi.unstubAllGlobals()
  useTabs.setState(initialTabs, true)
})

const typed: DocumentNode = {
  type: 'doc',
  content: [{ type: 'paragraph', content: [{ type: 'text', text: 'typed before the switch' }] }],
}

describe('a hidden tab', () => {
  it('autosaves what its editor showed when it was left, not what it opened with', async () => {
    const autosave = vi.fn(async (_request: { content: string }) => ({ ok: true, data: { savedAt: 1 } }))
    vi.stubGlobal('window', { api: { file: { autosave } } })
    const left = createWorkspaceStore()
    left.setState({ file: { path: null, name: 'a.sdoc', kind: DocumentKind.Document }, isDirty: true })
    addTab(left, null)
    const source = {
      readDoc: () => typed,
      readHtml: () => '',
      readPages: () => ({ pages: [], bands: [], sections: [] }),
    }
    left.getState().registerDocumentSource(source)

    addTab(createWorkspaceStore(), null)
    // As the editor does when it unmounts.
    left.getState().registerDocumentSource(null)
    await autosaveAllTabs()

    expect(autosave.mock.calls[0]?.[0].content).toContain('typed before the switch')
  })

  it('keeps its place and gets the screen back when activated', () => {
    const first = createWorkspaceStore()
    addTab(first, null)
    addTab(createWorkspaceStore(), null)
    const firstId = useTabs.getState().tabs[1]!.id

    activateTab(firstId)

    expect(useTabs.getState().activeId).toBe(firstId)
    expect(useTabs.getState().tabs).toHaveLength(3)
  })
})

describe('createTabSlot', () => {
  it('answers with the active tab value only', () => {
    useTabs.setState({ tabs: [...initialTabs.tabs, { id: 'tab-2', store: createWorkspaceStore() }] })
    const slot = createTabSlot<string>()
    slot.hold(initialTabs.activeId, 'first')

    useTabs.setState({ activeId: 'tab-2' })
    expect(slot.active()).toBeUndefined()

    slot.hold('tab-2', 'second')
    expect(slot.active()).toBe('second')
  })

  it('keeps a newer value when an older holder releases', () => {
    const slot = createTabSlot<string>()
    const releaseOld = slot.hold(initialTabs.activeId, 'old')
    const releaseNew = slot.hold(initialTabs.activeId, 'new')

    releaseOld()
    expect(slot.active()).toBe('new')

    releaseNew()
    expect(slot.active()).toBeUndefined()
  })
})
