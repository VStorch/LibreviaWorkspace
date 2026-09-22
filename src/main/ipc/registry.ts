import { ipcMain, type IpcMainInvokeEvent } from 'electron'
import type { InvocableIpcChannel } from '@shared/ipc-channels.js'
import { ipcContracts, type IpcRequest, type IpcResponse, type IpcResult } from '@shared/ipc.js'
import { AppError, ErrorCode, toSerializedError } from '@shared/errors.js'
import { t } from '../i18n.js'
import { editorPreferences } from '../preferences.js'

/**
 * Registro de handlers IPC.
 *
 * Toda payload vinda do renderer é validada contra o schema do canal antes de
 * chegar ao handler. O renderer é tratado como não confiável: se ele for
 * comprometido por um documento malicioso, não deve conseguir pedir uma
 * operação que o contrato não preveja.
 *
 * E a resposta é validada na volta, contra o `response` do mesmo contrato. Não
 * por desconfiança do main — por desconfiança do que ele **lê**: a lista de
 * fontes vem da saída de um programa do sistema, o conteúdo vem de disco, e o
 * schema é o único lugar onde os limites disso estão escritos. Enquanto ninguém
 * o executava, o contrato de resposta era tipo em tempo de compilação e mais
 * nada, e um handler fora de forma só aparecia na interface, longe da causa.
 *
 * Custa pouco: as respostas são rasas. Um documento de 50 MB atravessa como uma
 * `string`, e conferir uma `string` é conferir o tipo dela.
 */
export function handle<C extends InvocableIpcChannel>(
  channel: C,
  handler: (payload: IpcRequest<C>, event: IpcMainInvokeEvent) => Promise<IpcResponse<C>> | IpcResponse<C>,
): void {
  ipcMain.handle(channel, async (event, rawPayload: unknown): Promise<IpcResult<IpcResponse<C>>> => {
    try {
      const contract = ipcContracts[channel]
      const parsed = contract.request.safeParse(rawPayload)

      if (!parsed.success) {
        const field = parsed.error.issues[0]?.path.join('.')
        throw new AppError(
          ErrorCode.InvalidRequest,
          t('errors.ipc.invalidData'),
          field === undefined || field === '' ? undefined : `campo: ${field}`,
        )
      }

      const data = await handler(parsed.data as IpcRequest<C>, event)

      const replied = contract.response.safeParse(data)
      if (!replied.success) {
        const field = replied.error.issues[0]?.path.join('.')
        // Defeito nosso, não de quem usa: por isso `Internal`, e por isso o
        // detalhe técnico fica no `console.error` do bloco abaixo.
        throw new AppError(
          ErrorCode.Internal,
          t('errors.ipc.unrecognizedEnd'),
          field === undefined || field === '' ? undefined : `resposta, campo: ${field}`,
        )
      }

      // A resposta validada, e não a crua: o schema também normaliza — aplica
      // padrão e descarta campo que o contrato não declara.
      return { ok: true, data: replied.data as IpcResponse<C> }
    } catch (cause) {
      // O diagnóstico técnico fica aqui; o renderer recebe só o essencial.
      console.error(`[ipc] falha no canal ${channel}:`, cause)
      return { ok: false, error: toSerializedError(cause, editorPreferences().language) }
    }
  })
}
