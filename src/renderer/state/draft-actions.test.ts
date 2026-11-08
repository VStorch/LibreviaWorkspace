import { afterEach, describe, expect, it, vi } from 'vitest'
import { DocumentKind, type DraftSummary } from '@shared/types.js'
import { createWorkspaceStore } from './workspace-store.js'

const UUID = /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/

const DRAFT: DraftSummary = {
  draftId: '6f1c2b0e-8d4a-4c3e-9b7a-2e5d1f0a9c88',
  path: null,
  name: 'Sem título.sdoc',
  kind: DocumentKind.Document,
  savedAt: 1000,
}

function stubApi(drafts: readonly DraftSummary[] = []) {
  const api = {
    file: { autosave: vi.fn(async () => ({ ok: true, data: { savedAt: 1 } })) },
    recovery: {
      peek: vi.fn(async () => ({ ok: true, data: { drafts } })),
      discard: vi.fn(async () => ({ ok: true, data: { discarded: true } })),
    },
  }
  vi.stubGlobal('window', { api })
  return api
}

afterEach(() => {
  vi.unstubAllGlobals()
})

describe('draft actions', () => {
  it('gives each store its own draft id', () => {
    const first = createWorkspaceStore().getState().draftId
    const second = createWorkspaceStore().getState().draftId

    expect(first).toMatch(UUID)
    expect(second).toMatch(UUID)
    expect(first).not.toBe(second)
  })

  it('autosaves under the id of the store it belongs to', async () => {
    const api = stubApi()
    const store = createWorkspaceStore()
    store.setState({ file: { path: null, name: 'a.sdoc', kind: DocumentKind.Document }, isDirty: true })

    await store.getState().autosave()

    expect(api.file.autosave).toHaveBeenCalledWith(
      expect.objectContaining({ draftId: store.getState().draftId }),
    )
  })

  it('offers the most recent draft', async () => {
    stubApi([DRAFT, { ...DRAFT, draftId: '0b8e7f2a-1c3d-4e5f-8a9b-0c1d2e3f4a5b', savedAt: 500 }])
    const store = createWorkspaceStore()

    await store.getState().checkRecovery()

    expect(store.getState().pendingDraft).toEqual(DRAFT)
  })

  it('offers nothing when there is no draft', async () => {
    stubApi([])
    const store = createWorkspaceStore()

    await store.getState().checkRecovery()

    expect(store.getState().pendingDraft).toBeNull()
  })

  it('discards the offered draft, not the store own one', async () => {
    const api = stubApi()
    const store = createWorkspaceStore()
    store.setState({ pendingDraft: DRAFT })

    await store.getState().dismissDraft()

    expect(api.recovery.discard).toHaveBeenCalledWith({ draftId: DRAFT.draftId })
    expect(store.getState().pendingDraft).toBeNull()
  })
})
