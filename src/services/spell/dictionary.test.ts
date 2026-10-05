import { describe, expect, it } from 'vitest'
import { DICTIONARY_FOLDER, SPELL_LANGUAGE, dictionaryFileName, hasBdictSignature } from './dictionary.js'

describe('nome do dicionário', () => {
  it('monta o nome que o Chromium procura', () => {
    // The name is a contract, not a preference: with another name the spellchecker does not find
    // the bundled file and goes to download it, which on an offline machine means no checking and
    // no warning.
    expect(dictionaryFileName()).toBe('pt-BR-3-0.bdic')
    expect(dictionaryFileName('en-US')).toBe('en-US-3-0.bdic')
  })

  it('a pasta é a que o Chromium abre no diretório de dados', () => {
    expect(DICTIONARY_FOLDER).toBe('Dictionaries')
  })
})

describe('assinatura do formato binário', () => {
  it('reconhece um arquivo que começa em BDic', () => {
    expect(hasBdictSignature(new Uint8Array([0x42, 0x44, 0x69, 0x63, 0x02]))).toBe(true)
  })

  it('recusa arquivo curto ou de outro formato', () => {
    // A raw Hunspell (`.dic`) is text and would start with the word count: the easy mistake when
    // updating the dictionary by hand.
    expect(hasBdictSignature(new Uint8Array([0x42, 0x44]))).toBe(false)
    expect(hasBdictSignature(new TextEncoder().encode('123456\nabacate'))).toBe(false)
  })
})

describe('idioma', () => {
  it('verifica um idioma só, e é português do Brasil', () => {
    // A bilingual document is the rare case; offering a language list would be UI for a problem
    // this app does not have.
    expect(SPELL_LANGUAGE).toBe('pt-BR')
  })
})
