/**
 * A instância única do sidecar, já ligada ao ciclo de vida do Electron.
 *
 * Este é o único arquivo da pasta que conhece o `electron` — `client.ts` e
 * `locate.ts` ficam puros para serem testáveis sem subir o app.
 */

import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { app } from 'electron'
import { SidecarClient } from './client.js'
import { locateSidecarIn } from './locate.js'

/**
 * A raiz onde procurar `resources/sidecar/`.
 *
 * Fora do pacote, deriva da localização do próprio bundle (`out/main/index.js`
 * → dois níveis acima) em vez de `app.getAppPath()`. O `getAppPath()` muda
 * conforme o Electron é chamado — arquivo, pasta ou app empacotado — e essa
 * variação já custou um "instalação incompleta" com o binário no lugar certo.
 */
function resourceRoot(): string {
  if (app.isPackaged) return process.resourcesPath
  return join(dirname(fileURLToPath(import.meta.url)), '..', '..')
}

let instance: SidecarClient | null = null

export function sidecar(): SidecarClient {
  instance ??= new SidecarClient(() => locateSidecarIn(resourceRoot()))
  return instance
}

/**
 * Confere na subida se o serviço de formatos responde.
 *
 * **Não bloqueia a abertura do aplicativo.** Documento interno, texto e PDF não
 * passam pelo sidecar; travar tudo porque o serviço de DOCX não subiu seria
 * punir o usuário por um recurso que ele talvez nem vá usar. Quando ele de fato
 * abrir um `.docx`, aí sim recebe o erro — e aí ele é relevante.
 */
export async function checkSidecarHealth(): Promise<void> {
  try {
    const health = await sidecar().health()
    console.info(`[sidecar] ${health.name} ${health.version} — ${health.runtime}`)
  } catch (cause) {
    console.error('[sidecar] indisponível na subida:', cause instanceof Error ? cause.message : cause)
  }
}

/** Encerra o processo filho. Idempotente e seguro se nunca subiu. */
export function disposeSidecar(): void {
  instance?.dispose()
  instance = null
}
