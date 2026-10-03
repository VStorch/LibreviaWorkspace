/** O único arquivo da pasta que conhece o `electron`. */

import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { app } from 'electron'
import { SidecarClient } from './client.js'
import { locateSidecarIn } from './locate.js'

/** Fora do pacote, deriva do próprio bundle: `app.getAppPath()` muda conforme o Electron é chamado. */
function resourceRoot(): string {
  if (app.isPackaged) return process.resourcesPath
  return join(dirname(fileURLToPath(import.meta.url)), '..', '..')
}

let instance: SidecarClient | null = null

export function sidecar(): SidecarClient {
  instance ??= new SidecarClient(() => locateSidecarIn(resourceRoot()))
  return instance
}

/** Não bloqueia a abertura: documento interno, texto e PDF não passam pelo sidecar. */
export async function checkSidecarHealth(): Promise<void> {
  try {
    const health = await sidecar().health()
    console.info(`[sidecar] ${health.name} ${health.version} — ${health.runtime}`)
  } catch (cause) {
    console.error('[sidecar] indisponível na subida:', cause instanceof Error ? cause.message : cause)
  }
}

/** Idempotente, e seguro se nunca subiu. */
export function disposeSidecar(): void {
  instance?.dispose()
  instance = null
}
