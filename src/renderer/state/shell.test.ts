import { afterEach, describe, expect, it, vi } from 'vitest'
import { DocumentKind, type DraftSummary } from '@shared/types.js'
import { HOME_TAB_ID, MAX_OPEN_TABS } from '@services/tabs/tab-order.js'
import {
  checkRecovery,
  closeAllTabs,
  closeTab,
  dismissDraft,
  newDocument,
  openRecent,
  useShell,
} from './shell.js'
import { activeWorkspaceStore, addTab, useTabs } from './tabs.js'
import { createWorkspaceStore } from './workspace-store.js'

const initialTabs = useTabs.getState()

const DRAFT: DraftSummary = {
  draftId: '6f1c2b0e-8d4a-4c3e-9b7a-2e5d1f0a9c88',
  path: null,
  name: 'Sem título.sdoc',
  kind: DocumentKind.Document,
  savedAt: 1000,
}

const ok = <T>(data: T) => Promise.resolve({ ok: true as const, data })

function stubApi(drafts: readonly DraftSummary[] = []) {
  const api = {
    file: {
      openRecent: vi.fn(({ path }: { path: string }) =>
        ok({ file: { path, name: path.split('/').pop() ?? path, content: 'text' } }),
      ),
      release: vi.fn(() => ok({ released: true })),
    },
    recent: { list: vi.fn(() => ok({ files: [] })) },
    recovery: {
      peek: vi.fn(() => ok({ drafts })),
      discard: vi.fn(() => ok({ discarded: true })),
    },
    dialog: { confirmDiscard: vi.fn(() => ok({ choice: 'cancel' })) },
  }
  vi.stubGlobal('window', { api })
  return api
}

function openPaths(): (string | null)[] {
  return useTabs.getState().tabs.map((tab) => tab.store.getState().file?.path ?? null)
}

afterEach(() => {
  vi.unstubAllGlobals()
  useTabs.setState(initialTabs, true)
  useShell.setState({ pendingDraft: null, templateGallery: false })
})

describe('opening a file', () => {
  it('takes the place of Home', async () => {
    stubApi()

    await openRecent('/docs/a.txt')

    expect(openPaths()).toEqual(['/docs/a.txt'])
    expect(activeWorkspaceStore().getState().file?.path).toBe('/docs/a.txt')
  })

  it('opens each further file in a tab of its own', async () => {
    stubApi()

    await openRecent('/docs/a.txt')
    await openRecent('/docs/b.txt')

    expect(openPaths()).toEqual(['/docs/a.txt', '/docs/b.txt'])
  })

  it('shows the tab that already has the file instead of opening it twice', async () => {
    const api = stubApi()
    await openRecent('/docs/a.txt')
    await openRecent('/docs/b.txt')

    await openRecent('/docs//a.txt')

    expect(openPaths()).toEqual(['/docs/a.txt', '/docs/b.txt'])
    expect(activeWorkspaceStore().getState().file?.path).toBe('/docs/a.txt')
    expect(api.file.openRecent).toHaveBeenCalledTimes(2)
  })

  it('replaces an untouched blank document, as Word does', async () => {
    stubApi()
    newDocument()

    await openRecent('/docs/a.txt')

    expect(openPaths()).toEqual(['/docs/a.txt'])
  })

  it('keeps a blank document that was edited', async () => {
    stubApi()
    newDocument()
    activeWorkspaceStore().setState({ isDirty: true })

    await openRecent('/docs/a.txt')

    expect(openPaths()).toEqual([null, '/docs/a.txt'])
  })

  it('refuses beyond the limit, says why and asks main for nothing', async () => {
    const api = stubApi()
    for (let index = 0; index < MAX_OPEN_TABS; index += 1) {
      const store = createWorkspaceStore()
      store.setState({
        file: { path: `/docs/${index}.txt`, name: `${index}.txt`, kind: DocumentKind.Document },
      })
      addTab(store, null)
    }

    await openRecent('/docs/extra.txt')

    expect(api.file.openRecent).not.toHaveBeenCalled()
    expect(openPaths()).not.toContain('/docs/extra.txt')
    expect(activeWorkspaceStore().getState().error?.message).toContain(String(MAX_OPEN_TABS))
  })
})

describe('a new blank file', () => {
  it('opens beside an untouched blank one instead of replacing it', () => {
    stubApi()

    newDocument()
    newDocument()

    expect(useTabs.getState().tabs).toHaveLength(2)
  })
})

describe('closing a tab', () => {
  it('leaves Home alone after the last one', async () => {
    const api = stubApi()
    await openRecent('/docs/a.txt')

    expect(await closeTab(useTabs.getState().activeId)).toBe(true)

    expect(useTabs.getState().tabs.map((tab) => tab.id)).toEqual([HOME_TAB_ID])
    expect(api.file.release).toHaveBeenCalledWith({ path: '/docs/a.txt' })
  })

  it('keeps a modified tab when the user cancels', async () => {
    stubApi()
    await openRecent('/docs/a.txt')
    await openRecent('/docs/b.txt')
    const [first] = useTabs.getState().tabs
    first!.store.setState({ isDirty: true })

    expect(await closeTab(first!.id)).toBe(false)

    expect(openPaths()).toEqual(['/docs/a.txt', '/docs/b.txt'])
    expect(useTabs.getState().activeId).toBe(first!.id)
  })
})

describe('saving as', () => {
  it('refuses a path open in another tab, which would save over it', async () => {
    const api = stubApi()
    const save = vi.fn(() => ok({ name: 'a.txt' }))
    Object.assign(api.file, {
      chooseSavePath: vi.fn(() => ok({ canceled: false, path: '/docs/a.txt', name: 'a.txt' })),
      save,
    })
    await openRecent('/docs/a.txt')
    await openRecent('/docs/b.txt')
    const b = activeWorkspaceStore()

    expect(await b.getState().saveAs()).toBe(false)

    expect(save).not.toHaveBeenCalled()
    expect(b.getState().file?.path).toBe('/docs/b.txt')
    expect(b.getState().error?.message).toContain('a.txt')
  })
})

describe('closing all tabs', () => {
  it('closes every tab and leaves Home', async () => {
    stubApi()
    await openRecent('/docs/a.txt')
    await openRecent('/docs/b.txt')

    expect(await closeAllTabs()).toBe(true)

    expect(useTabs.getState().tabs.map((tab) => tab.id)).toEqual([HOME_TAB_ID])
  })

  it('stops at the first modified tab the user keeps', async () => {
    const api = stubApi()
    await openRecent('/docs/a.txt')
    await openRecent('/docs/b.txt')
    await openRecent('/docs/c.txt')
    useTabs.getState().tabs[1]!.store.setState({ isDirty: true })

    expect(await closeAllTabs()).toBe(false)

    expect(openPaths()).toEqual(['/docs/b.txt', '/docs/c.txt'])
    expect(api.dialog.confirmDiscard).toHaveBeenCalledTimes(1)
  })
})

describe('recovery', () => {
  it('offers the most recent draft', async () => {
    stubApi([DRAFT, { ...DRAFT, draftId: '0b8e7f2a-1c3d-4e5f-8a9b-0c1d2e3f4a5b', savedAt: 500 }])

    await checkRecovery()

    expect(useShell.getState().pendingDraft).toEqual(DRAFT)
  })

  it('offers nothing when there is no draft', async () => {
    stubApi([])

    await checkRecovery()

    expect(useShell.getState().pendingDraft).toBeNull()
  })

  it('discards the offered draft', async () => {
    const api = stubApi()
    useShell.setState({ pendingDraft: DRAFT })

    await dismissDraft()

    expect(api.recovery.discard).toHaveBeenCalledWith({ draftId: DRAFT.draftId })
    expect(useShell.getState().pendingDraft).toBeNull()
  })
})
