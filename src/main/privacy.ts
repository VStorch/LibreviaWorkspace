/**
 * The personal data inspector, for DOCX and XLSX: the sidecar reads and cleans the package as it is
 * on disk, and the clean copy becomes the opened file without being written. Saving writes it.
 */

import { readFile } from 'node:fs/promises'
import { z } from 'zod'
import { AppError, ErrorCode, fromFileSystemError } from '@shared/errors.js'
import { privacyReportSchema } from '@shared/schemas.js'
import { isExcelPath, isOfficePackagePath } from '@services/file/formats.js'
import type { PrivacyCategory } from '@services/file/privacy.js'
import { decodeDocx, type OpenedDocx } from './docx/index.js'
import { assertReadableFile } from './fs/paths.js'
import { t } from './i18n.js'
import { editorPreferences } from './preferences.js'
import type { SidecarClient } from './sidecar/client.js'
import { SidecarMethod } from './sidecar/protocol.js'
import { decodeXlsx, type OpenedXlsx } from './xlsx/index.js'

const scrubResultSchema = z.object({ changes: z.number().int().nonnegative() })

export async function inspectPersonalData(
  client: SidecarClient,
  path: string,
): Promise<z.infer<typeof privacyReportSchema>> {
  const reply = await client.request(
    SidecarMethod.PackageInspect,
    { anonymous: anonymousName() },
    await readPackage(path),
  )
  const parsed = privacyReportSchema.safeParse(reply.result)
  if (!parsed.success) {
    throw new AppError(
      ErrorCode.SidecarFailed,
      t('errors.privacy.cannotInspect'),
      t('errors.privacy.contract'),
    )
  }
  return parsed.data
}

export async function removePersonalData(
  client: SidecarClient,
  path: string,
  categories: readonly PrivacyCategory[],
): Promise<OpenedDocx | OpenedXlsx> {
  const reply = await client.request(
    SidecarMethod.PackageScrub,
    { anonymous: anonymousName(), categories },
    await readPackage(path),
  )
  if (!scrubResultSchema.safeParse(reply.result).success) {
    throw new AppError(
      ErrorCode.SidecarFailed,
      t('errors.privacy.cannotRemove'),
      t('errors.privacy.contract'),
    )
  }

  const clean = Buffer.from(reply.binary)
  return isExcelPath(path) ? decodeXlsx(client, path, clean) : decodeDocx(client, path, clean)
}

/** As Word names the authors it hides. */
function anonymousName(): string {
  return t('revisions.unknownAuthor')
}

/** From disk, not the opened original: after a save, that one is still the package saving started from. */
async function readPackage(path: string): Promise<Uint8Array> {
  if (!isOfficePackagePath(path)) {
    throw new AppError(ErrorCode.UnsupportedFormat, t('errors.privacy.onlyPackages'))
  }
  await assertReadableFile(path)
  try {
    return new Uint8Array(await readFile(path))
  } catch (cause) {
    throw fromFileSystemError(cause, 'read', editorPreferences().language)
  }
}
