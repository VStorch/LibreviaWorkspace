import { describe, expect, it } from 'vitest'
import {
  DEFAULT_NATURAL_LINE_HEIGHT,
  cssLineHeightOf,
  firstFontOf,
  lineFactorOf,
  naturalLineHeightOf,
} from './line-metrics.js'

describe('altura natural da linha', () => {
  it('a pilha do CSS é lida pela primeira família', () => {
    expect(firstFontOf('Calibri, sans-serif')).toBe('Calibri')
    expect(firstFontOf('"Times New Roman", serif')).toBe('Times New Roman')
    expect(firstFontOf(null)).toBeNull()
    expect(naturalLineHeightOf('Calibri, sans-serif')).toBe(1.2207)
  })

  it('sem fonte declarada vale a do editor; fonte de fora não tem palpite', () => {
    // As duas respostas vêm de `LineMetrics.Of`: silêncio é a nossa fonte, e
    // fonte que o instalador não leva cai numa substituta que depende da máquina.
    expect(naturalLineHeightOf(undefined)).toBe(DEFAULT_NATURAL_LINE_HEIGHT)
    expect(naturalLineHeightOf('Aptos')).toBeNull()
  })

  it('o fator do Word e o número do CSS fecham nos dois sentidos', () => {
    // 1,5 linha em Calibri são 1,83105 de CSS. O que o diálogo escolhe tem de
    // voltar igual, senão abrir e fechar o diálogo muda o documento.
    expect(cssLineHeightOf(1.5, 'Calibri, sans-serif')).toBe('1.8311')
    expect(lineFactorOf(1.8311, 'Calibri, sans-serif')).toBe(1.5)

    expect(cssLineHeightOf(1, 'Arial')).toBe('1.1499')
    expect(lineFactorOf(1.1499, 'Arial')).toBe(1)
  })

  it('fonte desconhecida com fator 1 sai como "normal"', () => {
    // É o que o leitor faz: sem saber a altura, quem mede melhor é o navegador.
    expect(cssLineHeightOf(1, 'Aptos')).toBe('normal')
    expect(cssLineHeightOf(2, 'Aptos')).toBe('2.2998')
    expect(lineFactorOf(2.2998, 'Aptos')).toBe(2)
  })
})
