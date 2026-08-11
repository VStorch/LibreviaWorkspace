import { app } from 'electron'
import { IpcChannel } from '@shared/ipc-channels.js'
import { handle } from './registry.js'

/**
 * Canal de fumaça da Fase 0.
 *
 * Existe para provar que o caminho renderer → preload → main → renderer está
 * completo, tipado e validado ponta a ponta. Será removido quando os canais
 * reais de arquivo entrarem, na Fase 1.
 */
export function registerAppHandlers(): void {
  handle(IpcChannel.AppPing, (payload) => ({
    echo: payload.message,
    receivedAt: Date.now(),
    versions: {
      app: app.getVersion(),
      electron: process.versions.electron ?? 'desconhecida',
      chrome: process.versions.chrome ?? 'desconhecida',
      node: process.versions.node ?? 'desconhecida',
    },
  }))
}
