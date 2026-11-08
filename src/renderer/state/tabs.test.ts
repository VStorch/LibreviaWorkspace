import { afterEach, describe, expect, it } from 'vitest'
import { createTabSlot, useTabs } from './tabs.js'
import { createWorkspaceStore } from './workspace-store.js'

const initialTabs = useTabs.getState()

afterEach(() => {
  useTabs.setState(initialTabs, true)
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
