import { ipcMain, type IpcMainInvokeEvent } from 'electron'
import type { IpcChannel } from '@shared/ipc-channels.js'
import { ipcContracts, type IpcRequest, type IpcResponse, type IpcResult } from '@shared/ipc.js'
import { AppError, ErrorCode, toSerializedError } from '@shared/errors.js'

/**
 * Registro de handlers IPC.
 *
 * Toda payload vinda do renderer é validada contra o schema do canal antes de
 * chegar ao handler. O renderer é tratado como não confiável: se ele for
 * comprometido por um documento malicioso, não deve conseguir pedir uma
 * operação que o contrato não preveja.
 */
export function handle<C extends IpcChannel>(
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
          'A operação foi recusada porque os dados enviados são inválidos.',
          field === undefined || field === '' ? undefined : `campo: ${field}`,
        )
      }

      const data = await handler(parsed.data as IpcRequest<C>, event)
      return { ok: true, data }
    } catch (cause) {
      // O diagnóstico técnico fica aqui; o renderer recebe só o essencial.
      console.error(`[ipc] falha no canal ${channel}:`, cause)
      return { ok: false, error: toSerializedError(cause) }
    }
  })
}
