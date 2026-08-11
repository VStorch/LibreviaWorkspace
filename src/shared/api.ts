import type { IpcChannel } from './ipc-channels.js'
import type { IpcRequest, IpcResponse, IpcResult } from './ipc.js'

/**
 * Superfície completa que o renderer enxerga do mundo externo.
 *
 * É deliberadamente uma interface explícita, e não um tipo inferido do preload:
 * ampliar o que o renderer pode fazer tem de ser uma edição consciente deste
 * arquivo, não um efeito colateral de mexer na implementação.
 */
export interface AppApi {
  readonly app: {
    ping(
      payload: IpcRequest<typeof IpcChannel.AppPing>,
    ): Promise<IpcResult<IpcResponse<typeof IpcChannel.AppPing>>>
  }
}
