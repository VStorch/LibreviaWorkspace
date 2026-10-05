import { closeSync, copyFileSync, mkdirSync, openSync, readFileSync, readSync } from 'node:fs'
import { readFile } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { app, protocol, type Session } from 'electron'
import {
  DICTIONARY_FOLDER,
  SPELL_LANGUAGE,
  dictionaryFileName,
  hasBdictSignature,
} from '@services/spell/dictionary.js'
import { DictionaryScope } from '@shared/types.js'

/**
 * Offline, even on first run. Chromium looks for the dictionary at `Dictionaries/pt-BR-3-0.bdic`
 * **before** downloading, so the bundled file is copied there.
 * `setSpellCheckerDictionaryDownloadURL` is still a download, and `setSpellCheckProvider` would
 * replace the whole spellchecker; the download address still points to a scheme of ours, as a
 * safety belt.
 */

export const DICTIONARY_SCHEME = 'librevia-dict'

/** Chromium has no ignore list: "ignore" is an entry removed on exit. */
const sessionWords = new Set<string>()

/** `app.getAppPath()` changes with how Electron is launched; see `src/main/fonts.ts`. */
function bundledDictionaryPath(): string {
  const root = app.isPackaged
    ? process.resourcesPath
    : join(dirname(fileURLToPath(import.meta.url)), '..', '..')
  return join(root, 'resources', 'dictionaries', dictionaryFileName())
}

function installedDictionaryPath(): string {
  return join(app.getPath('userData'), DICTIONARY_FOLDER, dictionaryFileName())
}

/**
 * Idempotent. Must run before the default session exists. `false` leaves spelling marking nothing,
 * better than keeping the app from opening.
 */
export function installBundledDictionary(): boolean {
  // On macOS the spellchecker is the system's, without `.bdic`.
  if (process.platform === 'darwin') return true

  const target = installedDictionaryPath()

  try {
    // Being there is not enough: Chromium deletes a broken `.bdic` and tries to download. Checking
    // the signature restores it in this run.
    if (hasInstalledSignature(target)) return true
    console.error(`[spellcheck] dicionário do perfil está corrompido e será reposto: ${target}`)
  } catch {
    // Does not exist yet: first run.
  }

  try {
    const source = bundledDictionaryPath()

    // Better to find out in the log than to be left without a spellchecker without knowing why.
    if (!hasBdictSignature(readFileSync(source))) {
      console.error(`[spellcheck] dicionário embutido não está no formato BDic: ${source}`)
      return false
    }

    mkdirSync(dirname(target), { recursive: true })
    copyFileSync(source, target)
    return true
  } catch (cause) {
    console.error('[spellcheck] não foi possível instalar o dicionário embutido:', cause)
    return false
  }
}

/**
 * Only the first four bytes: it runs synchronously on every start, and the dictionary has
 * megabytes.
 */
function hasInstalledSignature(target: string): boolean {
  const handle = openSync(target, 'r')
  try {
    const head = new Uint8Array(4)
    readSync(handle, head, 0, 4, 0)
    return hasBdictSignature(head)
  } finally {
    closeSync(handle)
  }
}

/**
 * Reaching here means the local file was not found: the warning goes to the log, and the same file
 * is served.
 */
export function serveDictionary(): void {
  protocol.handle(DICTIONARY_SCHEME, async (request) => {
    console.warn(`[spellcheck] o corretor pediu o dicionário pela rede: ${request.url}`)

    try {
      return new Response(await readFile(bundledDictionaryPath()), {
        headers: { 'Content-Type': 'application/octet-stream' },
      })
    } catch {
      return new Response('', { status: 404 })
    }
  })
}

/**
 * The download address is set **before** the language, because the language triggers the search for
 * the dictionary.
 */
export function applySpellChecker(session: Session, enabled: boolean): void {
  try {
    session.setSpellCheckerDictionaryDownloadURL(`${DICTIONARY_SCHEME}://dictionaries/`)

    if (enabled) session.setSpellCheckerLanguages([SPELL_LANGUAGE])
    session.setSpellCheckerEnabled(enabled)
  } catch (cause) {
    // Without a spellchecker in Electron (or on macOS without the language), spelling is lost, not
    // the app.
    console.error('[spellcheck] o corretor não pôde ser configurado:', cause)
  }
}

/** `session` is "ignore": removed from the dictionary on exit. */
export function rememberWord(session: Session, word: string, scope: DictionaryScope): boolean {
  const added = session.addWordToSpellCheckerDictionary(word)
  if (added && scope === DictionaryScope.Session) sessionWords.add(word)
  return added
}

export function forgetSessionWords(session: Session): void {
  for (const word of sessionWords) {
    try {
      session.removeWordFromSpellCheckerDictionary(word)
    } catch {
      // Failing to clean the dictionary must not delay shutdown.
    }
  }
  sessionWords.clear()
}
