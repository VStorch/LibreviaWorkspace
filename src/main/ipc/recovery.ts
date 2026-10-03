/** O rascunho é o que estava **na tela**, em formato interno mesmo quando a origem é `.docx` ou `.xlsx`. */

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
 * A autorização de gravação e os bytes originais morreram com o processo. Se o
 * arquivo sumiu, o trabalho recuperado vira um "salvar como".
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
