import { describe, expect, it } from 'vitest'
import {
  BUILTIN_STYLES,
  blockStyleOf,
  listedStyles,
  styleLabelOf,
  type StyleDefinition,
  type StyleSheet,
} from './styles.js'

/** Um documento do Word em português: o id é traduzido, o nome interno não. */
const portuguese: StyleSheet = {
  defaults: {
    paragraph: {},
    character: { fontFamily: 'Arial, sans-serif', fontSize: '11pt' },
    paragraphStyleId: 'Padro',
    characterStyleId: null,
  },
  styles: {
    Padro: { id: 'Padro', name: 'Normal', type: 'paragraph', qFormat: true, hidden: false, custom: false },
    Ttulo1: {
      id: 'Ttulo1',
      name: 'heading 1',
      type: 'paragraph',
      qFormat: true,
      hidden: false,
      custom: false,
      uiPriority: 9,
    },
    Citao: {
      id: 'Citao',
      name: 'Citação longa',
      type: 'paragraph',
      qFormat: true,
      hidden: false,
      custom: true,
      uiPriority: 1,
    },
    Oculto: {
      id: 'Oculto',
      name: 'Tabela sem nada',
      type: 'paragraph',
      qFormat: false,
      hidden: true,
      custom: false,
    },
  },
}

describe('o estilo do bloco onde está o cursor', () => {
  it('é o que o bloco trouxe do arquivo', () => {
    const style = blockStyleOf(portuguese, { type: 'paragraph', styleId: 'Citao' })
    expect(style?.name).toBe('Citação longa')
  })

  it('ignora um id que o documento não define', () => {
    // No Word um `w:pStyle` pendurado é silêncio: o parágrafo sai como o padrão.
    // Mostrar o id mesmo assim seria dizer que o documento tem um estilo que ele
    // não tem — e é o caso do `.sdoc` que veio de um `.docx` e voltou ao pacote
    // mínimo, que não define `Ttulo1`.
    const style = blockStyleOf(portuguese, { type: 'paragraph', styleId: 'Heading1' })
    expect(style?.id).toBe('Padro')
  })

  it('acha o título pelo nome interno, e não pelo id', () => {
    // O id é traduzido e o nome não: é o mesmo critério do leitor e do escritor
    // (`StyleResolver.HeadingLevelByName`). Pelo id, um título feito na tela
    // ficaria sem estilo em todo documento que não fosse em inglês.
    const style = blockStyleOf(portuguese, { type: 'heading', level: 1 })
    expect(style?.id).toBe('Ttulo1')
  })

  it('cai no estilo padrão do documento quando o bloco não diz nada', () => {
    const style = blockStyleOf(portuguese, { type: 'paragraph' })
    expect(style?.id).toBe('Padro')
  })

  it('não inventa resposta quando o documento não marca estilo padrão', () => {
    const sheet: StyleSheet = {
      defaults: { paragraph: {}, character: {}, paragraphStyleId: null, characterStyleId: null },
      styles: {},
    }
    expect(blockStyleOf(sheet, { type: 'paragraph' })).toBeNull()
  })

  it('num documento novo, o parágrafo é Normal e o título é o do nível', () => {
    expect(blockStyleOf(BUILTIN_STYLES, { type: 'paragraph' })?.id).toBe('Normal')
    expect(blockStyleOf(BUILTIN_STYLES, { type: 'heading', level: 3 })?.id).toBe('Heading3')
  })
})

describe('a lista que o painel mostra', () => {
  it('deixa de fora o que o documento esconde', () => {
    // `w:semiHidden` existe para tirar da lista a maquinaria do Word — a fonte
    // padrão do parágrafo e as dezenas de variantes de tabela que todo documento
    // declara sem usar. Com elas, o painel é uma lista que ninguém lê.
    expect(listedStyles(portuguese).map((style) => style.id)).toEqual(['Citao', 'Ttulo1', 'Padro'])
  })

  it('ordena pela prioridade que o documento declara', () => {
    // A ordem do Word: prioridade primeiro. `Padro` não declara nenhuma, e vai
    // para o fim — não para o começo, que é onde um `?? 0` o poria.
    const priorities = listedStyles(portuguese).map((style) => style.uiPriority ?? 100)
    expect(priorities).toEqual([...priorities].sort((left, right) => left - right))
  })

  it('não muda de ordem entre duas leituras do mesmo documento', () => {
    const twice = listedStyles(BUILTIN_STYLES).map((style) => style.id)
    expect(listedStyles(BUILTIN_STYLES).map((style) => style.id)).toEqual(twice)
  })
})

describe('o nome do estilo na tela', () => {
  it('traduz o nome interno dos títulos', () => {
    // `heading 1` é a forma que o arquivo exige para o Word reconhecer um título.
    // Mostrá-la seria mostrar o arquivo, e não o documento.
    const heading = BUILTIN_STYLES.styles['Heading1'] as StyleDefinition
    expect(styleLabelOf(heading)).toBe('Título 1')
  })

  it('mostra como está o nome que o autor escreveu', () => {
    const quotation = portuguese.styles['Citao'] as StyleDefinition
    expect(styleLabelOf(quotation)).toBe('Citação longa')
  })
})
