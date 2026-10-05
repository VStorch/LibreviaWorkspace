import { copyFile, open, rename, stat, unlink, type FileHandle } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { fromFileSystemError } from '@shared/errors.js'

/** Network folders sometimes do not implement fsync, and that is not a write failure. */
const FSYNC_UNSUPPORTED = new Set(['EINVAL', 'ENOTSUP', 'EPERM', 'EBADF', 'EISDIR'])

/**
 * No window for loss:
 *
 *  1. temp file **in the same folder**, because `rename()` across volumes fails with EXDEV, as on a
 *     network folder;
 *  2. fsync of the temp file;
 *  3. copy of the current file to `.bak` (skipped with `backup: false`, as for the draft rewritten
 *     every eight seconds);
 *  4. `rename` over the destination, the atomic swap;
 *  5. fsync of the folder, so the swap survives a crash.
 *
 * If anything fails, the temp file goes away and the original stays as it was.
 */
export async function writeFileAtomic(
  targetPath: string,
  data: string | Uint8Array,
  options: { backup?: boolean } = {},
): Promise<void> {
  const directory = dirname(targetPath)
  const temporaryPath = join(directory, `.${crypto.randomUUID()}.tmp`)

  let handle: FileHandle | undefined
  try {
    const existingMode = await modeOf(targetPath)

    // 'wx' fails if the temp file already exists, from another instance.
    handle = await open(temporaryPath, 'wx', existingMode ?? 0o666)
    await (typeof data === 'string' ? handle.writeFile(data, 'utf8') : handle.writeFile(data))
    await syncIfSupported(handle)
    await handle.close()
    handle = undefined

    if (existingMode !== null && options.backup !== false) {
      await copyFile(targetPath, `${targetPath}.bak`)
    }

    await rename(temporaryPath, targetPath)
    await syncDirectory(directory)
  } catch (cause) {
    if (handle !== undefined) await handle.close().catch(() => undefined)
    await unlink(temporaryPath).catch(() => undefined)
    throw fromFileSystemError(cause, 'escrita')
  }
}

/** So saving does not change permissions. */
async function modeOf(path: string): Promise<number | null> {
  try {
    return (await stat(path)).mode
  } catch {
    return null
  }
}

async function syncIfSupported(handle: FileHandle): Promise<void> {
  try {
    await handle.sync()
  } catch (cause) {
    const code = (cause as { code?: string }).code
    if (code === undefined || !FSYNC_UNSUPPORTED.has(code)) throw cause
  }
}

/** Best effort: Windows cannot open a directory for fsync, and no network guarantees it. */
async function syncDirectory(directory: string): Promise<void> {
  let handle: FileHandle | undefined
  try {
    handle = await open(directory, 'r')
    await handle.sync()
  } catch {
    // Best effort: without the folder fsync, the swap is already done.
  } finally {
    await handle?.close().catch(() => undefined)
  }
}
