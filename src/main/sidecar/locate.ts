/** Without `electron`: the root comes as a parameter, to be testable. */

import { access, constants } from 'node:fs/promises'
import { join } from 'node:path'
import { AppError, ErrorCode } from '@shared/errors.js'
import { t } from '../i18n.js'

export const SIDECAR_EXECUTABLE = 'Librevia.Format'

const unavailable = (): string => t('errors.sidecar.serviceNotFound')

/** Only the two published targets: any other pair reports no binary instead of looking for one. */
export function runtimeIdentifier(
  platform: NodeJS.Platform = process.platform,
  arch: string = process.arch,
): 'linux-x64' | 'win-x64' | null {
  if (arch !== 'x64') return null
  if (platform === 'linux') return 'linux-x64'
  if (platform === 'win32') return 'win-x64'
  return null
}

export function sidecarFileName(platform: NodeJS.Platform = process.platform): string {
  return platform === 'win32' ? `${SIDECAR_EXECUTABLE}.exe` : SIDECAR_EXECUTABLE
}

export function sidecarPathIn(root: string, platform: NodeJS.Platform = process.platform): string {
  const rid = runtimeIdentifier(platform)
  if (rid === null) {
    throw new AppError(
      ErrorCode.SidecarUnavailable,
      unavailable(),
      t('errors.sidecar.noBinary', { platform, arch: process.arch }),
    )
  }
  return join(root, 'resources', 'sidecar', rid, sidecarFileName(platform))
}

/** `LIBREVIA_SIDECAR_PATH` wins, so tests can use a fake sidecar. */
export async function locateSidecarIn(root: string): Promise<string> {
  const override = process.env['LIBREVIA_SIDECAR_PATH']
  const candidate = override !== undefined && override !== '' ? override : sidecarPathIn(root)

  try {
    await access(candidate, constants.X_OK)
  } catch {
    throw new AppError(ErrorCode.SidecarUnavailable, unavailable(), t('errors.sidecar.missingOrNotExecutable'))
  }

  return candidate
}
