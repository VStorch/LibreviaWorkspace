import { afterEach, describe, expect, it, vi } from 'vitest'
import { DocumentKind } from '@shared/types.js'
import { createWorkspaceStore } from './workspace-store.js'

const UUID = /^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$/

function stubApi() {
  const api = {
    file: { autosave: vi.fn(async () => ({ ok: true, data: { savedAt: 1 } })) },
    recovery: {
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
})
