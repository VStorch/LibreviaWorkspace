import { describe, expect, it } from 'vitest'
import {
  GUARANTEED_FONT_FAMILIES,
  familiesInDocument,
  firstFamilyOf,
  orderFontFamilies,
  parseFontconfigFamilies,
  parseWindowsFontRegistry,
} from './font-list.js'

describe('lista de fontes', () => {
  it('as do documento vêm primeiro', () => {
    // Quem abre um arquivo alheio procura a fonte **dele**. Enterrada no meio de
    // trezentas instaladas, a fonte do documento é tão inacessível quanto antes,
    // quando a lista tinha sete nomes fixos.
    const ordered = orderFontFamilies(['Arial', 'Zapfino'], ['Garamond'])

    expect(ordered[0]).toBe('Garamond')
  })

  it('as garantidas vêm antes das apenas instaladas', () => {
    const ordered = orderFontFamilies(['AAA Fonte', 'Calibri'])

    expect(ordered.slice(0, GUARANTEED_FONT_FAMILIES.length)).toEqual([...GUARANTEED_FONT_FAMILIES])
    expect(ordered.at(-1)).toBe('AAA Fonte')
  })

  it('lista vazia do sistema não esvazia o seletor', () => {
    // Sistema sem `fontconfig` devolve nada, e isso não é erro: a barra continua
    // oferecendo o que o instalador garante, que é o que ela oferecia antes.
    expect(orderFontFamilies([])).toEqual([...GUARANTEED_FONT_FAMILIES])
  })

  it('o mesmo nome não aparece duas vezes, nem com outra caixa', () => {
    const ordered = orderFontFamilies(['arial', 'ARIAL', 'Arial'], ['Arial'])

    expect(ordered.filter((family) => family.toLowerCase() === 'arial')).toHaveLength(1)
  })

  it('as instaladas saem em ordem alfabética', () => {
    // `fc-list` devolve na ordem do cache do fontconfig, que não é ordem nenhuma
    // para quem lê a lista.
    const ordered = orderFontFamilies(['Ubuntu', 'DejaVu Sans', 'Noto Serif'])
    const extras = ordered.filter((family) => !GUARANTEED_FONT_FAMILIES.includes(family))

    expect(extras).toEqual(['DejaVu Sans', 'Noto Serif', 'Ubuntu'])
  })

  it('a pilha de CSS do documento vira um nome', () => {
    expect(firstFamilyOf("'Times New Roman', Liberation Serif, serif")).toBe('Times New Roman')
    expect(firstFamilyOf('')).toBe('')
  })

  it('acha a fonte na marca do texto e no atributo do bloco', () => {
    // O leitor emite a fonte nos dois lugares: a altura da linha nasce da fonte
    // do elemento, não do que está escrito dentro dele.
    const doc = {
      type: 'doc',
      content: [
        {
          type: 'heading',
          attrs: { fontFamily: 'Cambria, Caladea, serif' },
          content: [{ type: 'text', marks: [{ type: 'textStyle', attrs: { fontFamily: 'Garamond' } }] }],
        },
      ],
    }

    expect(familiesInDocument(doc)).toEqual(['Cambria', 'Garamond'])
  })

  it('lê a saída do fontconfig', () => {
    const output = 'DejaVu Sans\nNimbus Sans,Nimbus Sans L\n\nUbuntu\nDejaVu Sans\n'

    expect(parseFontconfigFamilies(output)).toEqual(['DejaVu Sans', 'Nimbus Sans', 'Ubuntu'])
  })

  it('lê o registro do Windows sem repetir corte como família', () => {
    // O registro guarda um arquivo por corte. Sem cortar o sufixo, o seletor
    // ofereceria "Arial", "Arial Bold" e "Arial Italic" como três fontes.
    const output = [
      'HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Fonts',
      '    Arial & Arial Bold & Arial Italic (TrueType)    REG_SZ    arial.ttf',
      '    Arial Black (TrueType)    REG_SZ    ariblk.ttf',
      '    Segoe UI Light (TrueType)    REG_SZ    segoeuil.ttf',
      '',
    ].join('\r\n')

    expect(parseWindowsFontRegistry(output)).toEqual(['Arial', 'Arial Black', 'Segoe UI Light'])
  })
})
