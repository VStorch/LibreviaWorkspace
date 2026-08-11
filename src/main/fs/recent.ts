import { stat } from 'node:fs/promises'
import Store from 'electron-store'
import type { RecentFile } from '@shared/types.js'
import { fileNameFromPath, kindFromPath } from '@services/file/formats.js'
import { normalizePath } from './paths.js'

const MAX_RECENT = 10

interface RecentSchema {
  files: RecentFile[]
}

const store = new Store<RecentSchema>({
  name: 'recent-files',
  defaults: { files: [] },
  // Um JSON corrompido não pode impedir o aplicativo de abrir: a lista de
  // recentes é conveniência, não dado do usuário.
  clearInvalidConfig: true,
})

/**
 * Lista de recentes, já sem os arquivos que sumiram do disco.
 *
 * A poda acontece na leitura porque arquivos em pasta de rede desaparecem e
 * reaparecem o tempo todo — oferecer um item que só vai dar erro ao clicar é
 * pior do que não oferecer.
 */
export async function listRecentFiles(): Promise<readonly RecentFile[]> {
  const stored = store.get('files')
  const existing = await Promise.all(
    stored.map(async (file) => ((await isReadableFile(file.path)) ? file : null)),
  )
  const pruned = existing.filter((file): file is RecentFile => file !== null)

  if (pruned.length !== stored.length) store.set('files', pruned)
  return pruned
}

export function rememberRecentFile(path: string): void {
  const normalized = normalizePath(path)
  const entry: RecentFile = {
    path: normalized,
    name: fileNameFromPath(normalized),
    kind: kindFromPath(normalized),
    openedAt: Date.now(),
  }

  const withoutDuplicate = store.get('files').filter((file) => file.path !== normalized)
  store.set('files', [entry, ...withoutDuplicate].slice(0, MAX_RECENT))
}

export function clearRecentFiles(): void {
  store.set('files', [])
}

/** Um caminho só pode ser aberto por atalho se já estiver entre os recentes. */
export function isRemembered(path: string): boolean {
  const normalized = normalizePath(path)
  return store.get('files').some((file) => file.path === normalized)
}

async function isReadableFile(path: string): Promise<boolean> {
  try {
    return (await stat(path)).isFile()
  } catch {
    return false
  }
}
