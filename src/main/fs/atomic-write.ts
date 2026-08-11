import { copyFile, open, rename, stat, unlink, type FileHandle } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { fromFileSystemError } from '@shared/errors.js'

/** Sistemas de arquivos de rede às vezes não implementam fsync. Não é falha de gravação. */
const FSYNC_UNSUPPORTED = new Set(['EINVAL', 'ENOTSUP', 'EPERM', 'EBADF', 'EISDIR'])

/**
 * Grava um arquivo de texto sem janela de perda.
 *
 * A sequência importa, e cada passo existe por um motivo:
 *
 *  1. escreve num temporário **na mesma pasta** do destino — não em /tmp:
 *     `rename()` entre sistemas de arquivos diferentes falha com EXDEV, que é
 *     exatamente o caso de uma pasta de rede montada;
 *  2. faz fsync do temporário, para que os dados estejam no disco antes de
 *     qualquer coisa passar a apontar para eles;
 *  3. copia o arquivo atual para `.bak`, se já existir;
 *  4. renomeia o temporário sobre o destino — troca atômica no mesmo volume:
 *     em nenhum instante o destino fica truncado ou pela metade;
 *  5. faz fsync da pasta, para que a própria troca sobreviva a uma queda.
 *
 * Se qualquer passo falhar, o temporário é removido e o arquivo original
 * continua exatamente como estava.
 */
export async function writeTextFileAtomic(targetPath: string, content: string): Promise<void> {
  const directory = dirname(targetPath)
  const temporaryPath = join(directory, `.${crypto.randomUUID()}.tmp`)

  let handle: FileHandle | undefined
  try {
    const existingMode = await modeOf(targetPath)

    // 'wx' falha se o temporário já existir — evita colidir com outra instância.
    handle = await open(temporaryPath, 'wx', existingMode ?? 0o666)
    await handle.writeFile(content, 'utf8')
    await syncIfSupported(handle)
    await handle.close()
    handle = undefined

    if (existingMode !== null) {
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

/** Modo do arquivo existente, para que salvar não altere as permissões dele. */
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
    // Falta de suporte a fsync não invalida a gravação; qualquer outro erro sim.
    if (code === undefined || !FSYNC_UNSUPPORTED.has(code)) throw cause
  }
}

/**
 * Sincroniza a entrada de diretório. É melhor-esforço de propósito: o Windows
 * não permite abrir diretório para fsync, e nenhuma rede garante isso — mas
 * quando funciona, protege a troca de nome contra uma queda de energia.
 */
async function syncDirectory(directory: string): Promise<void> {
  let handle: FileHandle | undefined
  try {
    handle = await open(directory, 'r')
    await handle.sync()
  } catch {
    // silêncio proposital
  } finally {
    await handle?.close().catch(() => undefined)
  }
}
