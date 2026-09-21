import { describe, expect, it } from 'vitest'
import { DICTIONARY_FOLDER, SPELL_LANGUAGE, dictionaryFileName, hasBdictSignature } from './dictionary.js'

describe('nome do dicionário', () => {
  it('monta o nome que o Chromium procura', () => {
    // O nome é contrato, não preferência: com outro nome o corretor não acha o
    // arquivo embutido e parte para o download — que numa máquina offline
    // significa nenhuma verificação e nenhum aviso.
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
    // Um Hunspell cru (`.dic`) é texto e começaria com o número de palavras: é o
    // engano fácil de cometer ao atualizar o dicionário à mão.
    expect(hasBdictSignature(new Uint8Array([0x42, 0x44]))).toBe(false)
    expect(hasBdictSignature(new TextEncoder().encode('123456\nabacate'))).toBe(false)
  })
})

describe('idioma', () => {
  it('verifica um idioma só, e é português do Brasil', () => {
    // Um documento em duas línguas é o caso raro; oferecer uma lista de idiomas
    // para escolher seria interface para um problema que este aplicativo não tem.
    expect(SPELL_LANGUAGE).toBe('pt-BR')
  })
})
