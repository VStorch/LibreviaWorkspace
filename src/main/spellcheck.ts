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
 * Sem rede, nem na primeira execução. O Chromium procura o dicionário em
 * `Dictionaries/pt-BR-3-0.bdic` **antes** de baixar, e por isso o arquivo
 * embutido é copiado para lá. `setSpellCheckerDictionaryDownloadURL` continua a
 * ser download, e `setSpellCheckProvider` trocaria o corretor inteiro; o
 * endereço de download ainda aponta para um esquema nosso, como cinto de
 * segurança.
 */

export const DICTIONARY_SCHEME = 'librevia-dict'

/** O Chromium não tem lista de ignorados: "ignorar" é uma entrada desfeita ao sair. */
const sessionWords = new Set<string>()

/** `app.getAppPath()` muda conforme o Electron é chamado; ver `src/main/fonts.ts`. */
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
 * Idempotente. Precisa rodar antes de a sessão padrão existir. `false` deixa a
 * ortografia sem marcar nada, melhor que impedir o aplicativo de abrir.
 */
export function installBundledDictionary(): boolean {
  // No macOS o corretor é o do sistema, sem `.bdic`.
  if (process.platform === 'darwin') return true

  const target = installedDictionaryPath()

  try {
    // Estar lá não basta: um `.bdic` estragado o Chromium apaga e tenta baixar.
    // Conferindo a assinatura, ele é reposto nesta execução.
    if (hasInstalledSignature(target)) return true
    console.error(`[spellcheck] dicionário do perfil está corrompido e será reposto: ${target}`)
  } catch {
    // Ainda não existe: é a primeira execução.
  }

  try {
    const source = bundledDictionaryPath()

    // Melhor descobrir no log do que ficar sem corretor sem saber por quê.
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

/** Só os quatro primeiros bytes: roda a cada abertura, síncrono, e o dicionário tem megabytes. */
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

/** Chegar aqui é não ter achado o arquivo local: o aviso vai ao log, e o mesmo arquivo é servido. */
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
 * O endereço de download é definido **antes** do idioma, porque é o idioma que
 * dispara a procura pelo dicionário.
 */
export function applySpellChecker(session: Session, enabled: boolean): void {
  try {
    session.setSpellCheckerDictionaryDownloadURL(`${DICTIONARY_SCHEME}://dictionaries/`)

    if (enabled) session.setSpellCheckerLanguages([SPELL_LANGUAGE])
    session.setSpellCheckerEnabled(enabled)
  } catch (cause) {
    // Sem corretor no Electron (ou no macOS sem o idioma), perde-se a ortografia, e não o aplicativo.
    console.error('[spellcheck] o corretor não pôde ser configurado:', cause)
  }
}

/** `session` é o "ignorar": sai do dicionário ao encerrar. */
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
      // Falhar em limpar o dicionário não pode atrasar o encerramento.
    }
  }
  sessionWords.clear()
}
