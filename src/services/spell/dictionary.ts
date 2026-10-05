/**
 * The spellchecker is Chromium's, in main (`src/main/spellcheck.ts`). A file name other than the
 * one it looks for would make it try to download the dictionary, and offline there would be no
 * checking and no warning.
 */

export const SPELL_LANGUAGE = 'pt-BR'

/** The binary format revision, not the dictionary's: Chromium puts it in the name it looks for. */
export const DICTIONARY_REVISION = '3-0'

/** The folder Chromium opens inside the user data directory. */
export const DICTIONARY_FOLDER = 'Dictionaries'

/** Chromium looks for `<language>-<revision>.bdic` in the local folder before downloading. */
export function dictionaryFileName(language: string = SPELL_LANGUAGE): string {
  return `${language}-${DICTIONARY_REVISION}.bdic`
}

/**
 * Chromium deletes an invalid `.bdic` and tries to download: checking it in a test turns that into
 * a build failure.
 */
export function hasBdictSignature(bytes: Uint8Array): boolean {
  if (bytes.length < 4) return false
  return bytes[0] === 0x42 && bytes[1] === 0x44 && bytes[2] === 0x69 && bytes[3] === 0x63
}
