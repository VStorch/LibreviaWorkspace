/**
 * Onde está o binário do sidecar.
 *
 * Sem `import { app } from 'electron'` de propósito — a raiz chega como
 * parâmetro. É a mesma disciplina de `security-policy.ts` e `fs/paths.ts`:
 * lógica que dá para testar sem subir o Electron inteiro é lógica que vai ser
 * testada de verdade.
 */

import { access, constants } from 'node:fs/promises'
import { join } from 'node:path'
import { AppError, ErrorCode } from '@shared/errors.js'
import { t } from '../i18n.js'

export const SIDECAR_EXECUTABLE = 'Librevia.Format'

const unavailable = (): string => t('errors.sidecar.serviceNotFound')

/**
 * Identificador de runtime do .NET.
 *
 * São os dois alvos decididos em §8.1 do plano. Qualquer outro par
 * plataforma/arquitetura não tem binário publicado, e dizer isso é melhor do
 * que procurar um arquivo que nunca existiu.
 */
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

/** Caminho esperado dentro de uma raiz, sem tocar no disco. */
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

/**
 * Resolve o executável, já verificado como existente e executável.
 *
 * `LIBREVIA_SIDECAR_PATH` tem prioridade para que os testes apontem para um
 * sidecar de mentira sem precisar publicar o projeto .NET.
 */
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
