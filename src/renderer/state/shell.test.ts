import { afterEach, describe, expect, it, vi } from 'vitest'
import { DocumentKind, type DraftSummary } from '@shared/types.js'
import { createEmptyDocument } from '@services/document/model.js'
import { serializeDocument } from '@services/document/serialize.js'
import { HOME_TAB_ID, MAX_OPEN_TABS } from '@services/tabs/tab-order.js'
import {
  checkRecovery,
  closeAllTabs,
  closeTab,
  dismissDraft,
  newDocument,
  newFromTemplate,
  openRecent,
  openViaDialog,
  recoverDraft,
  saveAllTabs,
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

const EMPTY = serializeDocument(createEmptyDocument())

const ok = <T>(data: T) => Promise.resolve({ ok: true as const, data })

function stubApi(drafts: readonly DraftSummary[] = []) {
  const api = {
    file: {
      openRecent: vi.fn(({ path }: { path: string }) =>
        ok({ file: { path, name: path.split('/').pop() ?? path, content: 'text' } }),
      ),
      release: vi.fn(() => ok({ released: true })),
      open: vi.fn(),
      chooseSavePath: vi.fn(),
      save: vi.fn(),
    },
    recent: { list: vi.fn(() => ok({ files: [] })) },
    template: {
      open: vi.fn(() =>
        ok({ file: { path: '/models/carta.dotx', name: 'Carta.docx', content: EMPTY, template: true } }),
      ),
    },
    recovery: {
      peek: vi.fn(() => ok({ drafts })),
      restore: vi.fn(() => ok({ draft: null })),
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

describe('the original main keeps for each file', () => {
  const template = { source: 'builtin' as const, id: 'carta' }

  it('stays while another tab made from the same template still needs it', async () => {
    const api = stubApi()
    await newFromTemplate(template)
    await newFromTemplate(template)
    const [first, second] = useTabs.getState().tabs

    await closeTab(first!.id)
    expect(api.file.release).not.toHaveBeenCalledWith({ path: '/models/carta.dotx' })

    await closeTab(second!.id)
    expect(api.file.release).toHaveBeenCalledWith({ path: '/models/carta.dotx' })
  })

  it('leaves with the template a document saved as elsewhere no longer needs', async () => {
    const api = stubApi()
    Object.assign(api.file, {
      chooseSavePath: vi.fn(() => ok({ canceled: false, path: '/docs/carta.docx', name: 'carta.docx' })),
      save: vi.fn(() => ok({ name: 'carta.docx' })),
    })
    await newFromTemplate(template)

    expect(await activeWorkspaceStore().getState().saveAs()).toBe(true)

    expect(api.file.save).toHaveBeenCalledWith(expect.objectContaining({ origin: '/models/carta.dotx' }))
    expect(api.file.release).toHaveBeenCalledWith({ path: '/models/carta.dotx' })
  })

  it('is not read again when the dialog picks a file already open: its tab is shown', async () => {
    const api = stubApi()
    const open = vi.fn(() => ok({ canceled: false, alreadyOpen: '/docs/a.txt' }))
    Object.assign(api.file, { open })
    await openRecent('/docs/a.txt')
    await openRecent('/docs/b.txt')

    await openViaDialog()

    expect(open).toHaveBeenCalledWith({ openPaths: ['/docs/a.txt', '/docs/b.txt'] })
    expect(openPaths()).toEqual(['/docs/a.txt', '/docs/b.txt'])
    expect(activeWorkspaceStore().getState().file?.path).toBe('/docs/a.txt')
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

describe('saving every tab before quitting', () => {
  it('saves the modified tabs in order without showing those that have a path', async () => {
    const api = stubApi()
    const save = vi.fn(({ path }: { path: string }) => ok({ name: path.split('/').pop() }))
    Object.assign(api.file, { save })
    await openRecent('/docs/a.txt')
    await openRecent('/docs/b.txt')
    await openRecent('/docs/c.txt')
    const [a, , c] = useTabs.getState().tabs
    a!.store.setState({ isDirty: true })
    c!.store.setState({ isDirty: true })
    useTabs.setState({ activeId: useTabs.getState().tabs[1]!.id })

    expect(await saveAllTabs()).toBe(true)

    expect(save.mock.calls.map(([request]) => request.path)).toEqual(['/docs/a.txt', '/docs/c.txt'])
    expect(activeWorkspaceStore().getState().file?.path).toBe('/docs/b.txt')
  })

  it('stops at the first save that fails and shows its tab', async () => {
    const api = stubApi()
    const save = vi.fn(() =>
      Promise.resolve({ ok: false as const, error: { code: 'INTERNAL', message: 'x' } }),
    )
    Object.assign(api.file, { save })
    await openRecent('/docs/a.txt')
    await openRecent('/docs/b.txt')
    for (const tab of useTabs.getState().tabs) tab.store.setState({ isDirty: true })

    expect(await saveAllTabs()).toBe(false)

    expect(save).toHaveBeenCalledTimes(1)
    expect(activeWorkspaceStore().getState().file?.path).toBe('/docs/a.txt')
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

  it('does not offer the draft of a tab open now', async () => {
    const older = { ...DRAFT, draftId: '0b8e7f2a-1c3d-4e5f-8a9b-0c1d2e3f4a5b', savedAt: 500 }
    stubApi([DRAFT, older])
    newDocument()
    activeWorkspaceStore().setState({ draftId: DRAFT.draftId })

    await checkRecovery()

    expect(useShell.getState().pendingDraft).toEqual(older)
  })

  it('offers the next draft after one is recovered or discarded', async () => {
    const older = { ...DRAFT, draftId: '0b8e7f2a-1c3d-4e5f-8a9b-0c1d2e3f4a5b', savedAt: 500 }
    const api = stubApi([older])
    useShell.setState({ pendingDraft: DRAFT })

    await dismissDraft()

    expect(api.recovery.peek).toHaveBeenCalled()
    expect(useShell.getState().pendingDraft).toEqual(older)
  })

  it('does not recover the draft of a file open in a tab, which would hold it twice', async () => {
    const api = stubApi()
    await openRecent('/docs/a.txt')
    newDocument()
    useShell.setState({ pendingDraft: { ...DRAFT, path: '/docs/a.txt', name: 'a.txt' } })

    await recoverDraft()

    expect(api.recovery.restore).not.toHaveBeenCalled()
    expect(useShell.getState().pendingDraft?.path).toBe('/docs/a.txt')
    expect(activeWorkspaceStore().getState().file?.path).toBe('/docs/a.txt')
    expect(activeWorkspaceStore().getState().error?.message).toContain('a.txt')
  })

  it('discards the offered draft', async () => {
    const api = stubApi()
    useShell.setState({ pendingDraft: DRAFT })

    await dismissDraft()

    expect(api.recovery.discard).toHaveBeenCalledWith({ draftId: DRAFT.draftId })
    expect(useShell.getState().pendingDraft).toBeNull()
  })
})
