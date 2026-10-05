/**
 * The draft is what was **on screen**, in the internal format even when the source is `.docx` or
 * `.xlsx`.
 */

import { stat } from 'node:fs/promises'
import { IpcChannel } from '@shared/ipc-channels.js'
import { isExcelPath, isWordPackagePath } from '@services/file/formats.js'
import { adoptDocxOriginal } from '../docx/index.js'
import { adoptXlsxOriginal } from '../xlsx/index.js'
import { authorizePath } from '../fs/paths.js'
import { discardDraft, readDraft, readDraftSummary, writeDraft } from '../fs/recovery.js'
import { handle } from './registry.js'

export function registerRecoveryHandlers(): void {
  handle(IpcChannel.FileAutosave, async (payload) => ({
    savedAt: await writeDraft({
      path: payload.path,
      name: payload.name,
      kind: payload.kind,
      content: payload.content,
    }),
  }))

  handle(IpcChannel.RecoveryPeek, async () => ({ draft: await readDraftSummary() }))

  handle(IpcChannel.RecoveryRestore, async () => {
    const draft = await readDraft()
    if (draft === null) return { draft: null }

    if (draft.path !== null) await reattach(draft.path)
    return { draft }
  })

  handle(IpcChannel.RecoveryDiscard, async () => {
    await discardDraft()
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
  if (isWordPackagePath(path)) await adoptDocxOriginal(path)
  if (isExcelPath(path)) await adoptXlsxOriginal(path)
}

async function exists(path: string): Promise<boolean> {
  try {
    return (await stat(path)).isFile()
  } catch {
    return false
  }
}
