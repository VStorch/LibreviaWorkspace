/**
 * The draft is what was **on screen**, in the internal format even when the source is `.docx` or
 * `.xlsx`.
 */

import { stat } from 'node:fs/promises'
import { BrowserWindow } from 'electron'
import type { IpcRequest } from '@shared/ipc.js'
import { IpcChannel } from '@shared/ipc-channels.js'
import { isExcelPath, isOdtPackagePath, isWordPackagePath } from '@services/file/formats.js'
import { adoptOriginal, OriginalFormat } from '../opened-original.js'
import { authorizePath } from '../fs/paths.js'
import { discardDraft, listDraftSummaries, readDraft, writeDraft } from '../fs/recovery.js'
import { handle } from './registry.js'
import { windowOf } from './sender-window.js'

/** Which window writes each draft: another live window's drafts are its tabs, not left by a crash. */
const draftWindows = new Map<string, number>()

function heldByAnotherWindow(draftId: string, asking: BrowserWindow): boolean {
  const owner = draftWindows.get(draftId)
  if (owner === undefined || owner === asking.id) return false
  const window = BrowserWindow.fromId(owner)
  return window !== null && !window.isDestroyed()
}

export function registerRecoveryHandlers(): void {
  handle(IpcChannel.FileAutosave, async (payload, event) => {
    draftWindows.set(payload.draftId, windowOf(event).id)
    return { savedAt: await writeDraftOf(payload) }
  })

  handle(IpcChannel.RecoveryPeek, async (_payload, event) => {
    const asking = windowOf(event)
    const drafts = await listDraftSummaries()
    return { drafts: drafts.filter((draft) => !heldByAnotherWindow(draft.draftId, asking)) }
  })

  handle(IpcChannel.RecoveryRestore, async ({ draftId }) => {
    const draft = await readDraft(draftId)
    if (draft === null) return { draft: null }

    if (draft.path !== null) await reattach(draft.path)
    return { draft }
  })

  handle(IpcChannel.RecoveryDiscard, async ({ draftId }) => {
    await discardDraft(draftId)
    return { discarded: true as const }
  })
}

/**
 * The write authorization and the original bytes died with the process. If the file vanished, the
 * recovered work becomes a "save as".
 */
async function reattach(path: string): Promise<void> {
  if (!(await exists(path))) return

  authorizePath(path)
  if (isWordPackagePath(path)) await adoptOriginal(path, OriginalFormat.Docx)
  if (isOdtPackagePath(path)) await adoptOriginal(path, OriginalFormat.Odt)
  if (isExcelPath(path)) await adoptOriginal(path, OriginalFormat.Xlsx)
}

async function exists(path: string): Promise<boolean> {
  try {
    return (await stat(path)).isFile()
  } catch {
    return false
  }
}

function writeDraftOf(payload: IpcRequest<typeof IpcChannel.FileAutosave>): Promise<number> {
  return writeDraft(payload.draftId, {
    path: payload.path,
    name: payload.name,
    kind: payload.kind,
    content: payload.content,
  })
}
