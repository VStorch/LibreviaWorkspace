/**
 * Autosave e recuperação.
 *
 * O rascunho não é uma segunda cópia do arquivo: é o que estava **na tela**. Por
 * isso ele guarda o formato interno (`.sdoc`/`.ssheet`) mesmo quando o arquivo
 * de origem é `.docx` ou `.xlsx` — recuperar tem de devolver o documento como
 * ele estava sendo editado, e não como estava no disco antes da queda.
 */

import { stat } from 'node:fs/promises'
import { IpcChannel } from '@shared/ipc-channels.js'
import { isExcelPath, isWordPath } from '@services/file/formats.js'
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
 * Devolve ao caminho recuperado o que a sessão anterior sabia sobre ele.
 *
 * Duas coisas morreram com o processo: a autorização de gravação e os bytes
 * originais do pacote OOXML. Sem a primeira, salvar seria recusado; sem os
 * segundos, a gravação cirúrgica não teria sobre o que operar.
 *
 * O arquivo pode ter sumido no meio tempo — foi apagado, ou estava numa pasta de
 * rede que caiu junto. Nesse caso não há o que reatar, e o trabalho recuperado
 * continua válido: ele vira um "salvar como".
 */
async function reattach(path: string): Promise<void> {
  if (!(await exists(path))) return

  authorizePath(path)
  if (isWordPath(path)) await adoptDocxOriginal(path)
  if (isExcelPath(path)) await adoptXlsxOriginal(path)
}

async function exists(path: string): Promise<boolean> {
  try {
    return (await stat(path)).isFile()
  } catch {
    return false
  }
}
