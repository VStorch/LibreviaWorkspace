/** The only file in this folder that knows `electron`. */

import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { app } from 'electron'
import { SidecarClient } from './client.js'
import { locateSidecarIn } from './locate.js'

/**
 * Outside the package it derives from the bundle: `app.getAppPath()` changes with how Electron is
 * launched.
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

/** Does not block opening: internal documents, text and PDF do not go through the sidecar. */
export async function checkSidecarHealth(): Promise<void> {
  try {
    const health = await sidecar().health()
    console.info(`[sidecar] ${health.name} ${health.version} — ${health.runtime}`)
  } catch (cause) {
    console.error('[sidecar] indisponível na subida:', cause instanceof Error ? cause.message : cause)
  }
}

/** Idempotent, and safe if it never started. */
export function disposeSidecar(): void {
  instance?.dispose()
  instance = null
}
