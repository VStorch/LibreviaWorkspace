/**
 * O corretor é o do Chromium, no main (`src/main/spellcheck.ts`). Um nome de
 * arquivo diferente do que ele procura o faria tentar baixar o dicionário, e
 * offline não haveria verificação nem aviso.
 */

export const SPELL_LANGUAGE = 'pt-BR'

/** A revisão do formato binário, e não do dicionário: o Chromium a põe no nome que procura. */
export const DICTIONARY_REVISION = '3-0'

/** Pasta que o Chromium abre dentro do diretório de dados do usuário. */
export const DICTIONARY_FOLDER = 'Dictionaries'

/** O Chromium procura `<idioma>-<revisão>.bdic` na pasta local antes de baixar. */
export function dictionaryFileName(language: string = SPELL_LANGUAGE): string {
  return `${language}-${DICTIONARY_REVISION}.bdic`
}

/** Um `.bdic` inválido o Chromium apaga e tenta baixar: conferir em teste vira falha de build. */
export function hasBdictSignature(bytes: Uint8Array): boolean {
  if (bytes.length < 4) return false
  return bytes[0] === 0x42 && bytes[1] === 0x44 && bytes[2] === 0x69 && bytes[3] === 0x63
}
