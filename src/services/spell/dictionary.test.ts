import { describe, expect, it } from 'vitest'
import { DICTIONARY_FOLDER, SPELL_LANGUAGE, dictionaryFileName, hasBdictSignature } from './dictionary.js'

describe('dictionary name', () => {
  it('builds the name Chromium looks for', () => {
    // The name is a contract, not a preference: with another name the spellchecker does not find
    // the bundled file and goes to download it, which on an offline machine means no checking and
    // no warning.
    expect(dictionaryFileName()).toBe('pt-BR-3-0.bdic')
    expect(dictionaryFileName('en-US')).toBe('en-US-3-0.bdic')
  })

  it('the folder is the one Chromium opens in the data directory', () => {
    expect(DICTIONARY_FOLDER).toBe('Dictionaries')
  })
})

describe('binary format signature', () => {
  it('recognizes a file that starts with BDic', () => {
    expect(hasBdictSignature(new Uint8Array([0x42, 0x44, 0x69, 0x63, 0x02]))).toBe(true)
  })

  it('rejects a file that is short or in another format', () => {
    // A raw Hunspell (`.dic`) is text and would start with the word count: the easy mistake when
    // updating the dictionary by hand.
    expect(hasBdictSignature(new Uint8Array([0x42, 0x44]))).toBe(false)
    expect(hasBdictSignature(new TextEncoder().encode('123456\nabacate'))).toBe(false)
  })
})

describe('language', () => {
  it('checks a single language, Brazilian Portuguese', () => {
    // A bilingual document is the rare case; offering a language list would be UI for a problem
    // this app does not have.
    expect(SPELL_LANGUAGE).toBe('pt-BR')
  })
})
