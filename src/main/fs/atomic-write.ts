import { copyFile, open, rename, stat, unlink, type FileHandle } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { fromFileSystemError } from '@shared/errors.js'

/** Pastas de rede às vezes não implementam fsync, e isso não é falha de gravação. */
const FSYNC_UNSUPPORTED = new Set(['EINVAL', 'ENOTSUP', 'EPERM', 'EBADF', 'EISDIR'])

/**
 * Sem janela de perda:
 *
 *  1. temporário **na mesma pasta**, porque `rename()` entre volumes falha com
 *     EXDEV, como numa pasta de rede;
 *  2. fsync do temporário;
 *  3. cópia do atual para `.bak` (dispensável com `backup: false`, como no
 *     rascunho, reescrito a cada oito segundos);
 *  4. `rename` sobre o destino, a troca atômica;
 *  5. fsync da pasta, para a troca sobreviver a uma queda.
 *
 * Se algo falhar, o temporário sai e o original fica como estava.
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

    // 'wx' falha se o temporário já existir, de outra instância.
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

/** Para salvar não alterar as permissões. */
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

/** Melhor esforço: o Windows não abre diretório para fsync, e rede nenhuma garante. */
async function syncDirectory(directory: string): Promise<void> {
  let handle: FileHandle | undefined
  try {
    handle = await open(directory, 'r')
    await handle.sync()
  } catch {
    // Melhor esforço: sem fsync da pasta, a troca já está feita.
  } finally {
    await handle?.close().catch(() => undefined)
  }
}
