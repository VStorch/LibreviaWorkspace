import { ipcMain, type IpcMainInvokeEvent } from 'electron'
import type { InvocableIpcChannel } from '@shared/ipc-channels.js'
import { ipcContracts, type IpcRequest, type IpcResponse, type IpcResult } from '@shared/ipc.js'
import { AppError, ErrorCode, toSerializedError } from '@shared/errors.js'
import { t } from '../i18n.js'
import { editorPreferences } from '../preferences.js'

/**
 * The request is validated against the channel schema: a renderer compromised by a document cannot
 * ask for an operation the contract does not allow. The response too, because main reads disk and
 * program output, and the schema is where the limits are.
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
        // Our defect, not the user's: `Internal`, with the detail in the log.
        throw new AppError(
          ErrorCode.Internal,
          t('errors.ipc.unrecognizedEnd'),
          field === undefined || field === '' ? undefined : `resposta, campo: ${field}`,
        )
      }

      // The validated value, not the raw one: the schema applies defaults and drops undeclared
      // fields.
      return { ok: true, data: replied.data as IpcResponse<C> }
    } catch (cause) {
      console.error(`[ipc] falha no canal ${channel}:`, cause)
      return { ok: false, error: toSerializedError(cause, editorPreferences().language) }
    }
  })
}
