import type { Catalog } from '../message.js'

/**
 * Referências: navegação, marcadores, sumário, legendas e referências cruzadas.
 *
 * Área própria porque o M8 inteiro mora aqui, e parte do que ela guarda não é
 * interface e sim **texto do documento** — "Sumário", "Figura", "Tabela" — que
 * entra no arquivo no idioma de quem o escreve.
 */
export const REFERENCES = {
  'references.nav.title': { pt: 'Navegação', en: 'Navigation' },
  'references.nav.empty': {
    pt: 'Este documento não tem títulos. Aplique um estilo de título para vê-los aqui.',
    en: 'This document has no headings. Apply a heading style to see them here.',
  },
  'references.nav.entry': { pt: 'Nível {level}: {text}', en: 'Level {level}: {text}' },

  'references.bookmark.title': { pt: 'Marcadores', en: 'Bookmarks' },
  'references.bookmark.name': { pt: 'Nome do marcador', en: 'Bookmark name' },
  'references.bookmark.invalid': {
    pt: 'O nome começa por letra e só leva letras, algarismos e sublinhado, sem espaço.',
    en: 'The name starts with a letter and has only letters, digits and underscores, no spaces.',
  },
  'references.bookmark.list': { pt: 'Marcadores do documento', en: 'Bookmarks in the document' },
  'references.bookmark.empty': { pt: 'Nenhum marcador.', en: 'No bookmarks.' },
  'references.bookmark.order': { pt: 'Classificar por', en: 'Sort by' },
  'references.bookmark.byName': { pt: 'Nome', en: 'Name' },
  'references.bookmark.byLocation': { pt: 'Local', en: 'Location' },
  'references.bookmark.hidden': { pt: 'Marcadores ocultos', en: 'Hidden bookmarks' },
  'references.bookmark.readOnly': {
    pt: 'Documento em somente leitura: só é possível ir para um marcador.',
    en: 'Read-only document: you can only go to a bookmark.',
  },
  'references.bookmark.add': { pt: 'Adicionar', en: 'Add' },
  'references.bookmark.move': { pt: 'Mover para cá', en: 'Move here' },
  'references.bookmark.delete': { pt: 'Excluir', en: 'Delete' },
  'references.bookmark.goTo': { pt: 'Ir para', en: 'Go to' },

  'references.link.place': { pt: 'Lugar neste documento', en: 'Place in this document' },
  'references.link.noPlace': { pt: '(nenhum — usar o endereço)', en: '(none — use the address)' },
  'references.link.headings': { pt: 'Títulos', en: 'Headings' },
  'references.link.bookmarks': { pt: 'Marcadores', en: 'Bookmarks' },

  // Texto do documento, e não da interface: entra no arquivo no idioma de quem
  // escreve, como o Word faz.
  'references.toc.title': { pt: 'Sumário', en: 'Contents' },
  'references.toc.empty': {
    pt: 'Nenhuma entrada de sumário foi encontrada.',
    en: 'No table of contents entries found.',
  },
  'references.field.missingBookmark': {
    pt: 'Erro! Indicador não definido.',
    en: 'Error! Bookmark not defined.',
  },

  'references.insert': { pt: 'Inserir', en: 'Insert' },
  'references.caption.title': { pt: 'Legenda', en: 'Caption' },
  'references.caption.label': { pt: 'Rótulo', en: 'Label' },
  'references.caption.text': { pt: 'Texto depois do número', en: 'Text after the number' },
  'references.caption.position': { pt: 'Posição', en: 'Position' },
  'references.caption.below': { pt: 'Abaixo do item do cursor', en: 'Below the current item' },
  'references.caption.above': { pt: 'Acima do item do cursor', en: 'Above the current item' },
  // Rótulos do documento, como os do Word no idioma de quem escreve.
  'references.caption.figure': { pt: 'Figura', en: 'Figure' },
  'references.caption.table': { pt: 'Tabela', en: 'Table' },
  'references.caption.equation': { pt: 'Equação', en: 'Equation' },
  'references.crossRef.title': { pt: 'Referência cruzada', en: 'Cross-reference' },
  'references.crossRef.type': { pt: 'Tipo', en: 'Reference type' },
  'references.crossRef.heading': { pt: 'Título', en: 'Heading' },
  'references.crossRef.bookmark': { pt: 'Marcador', en: 'Bookmark' },
  'references.crossRef.target': { pt: 'Para qual', en: 'For which' },
  'references.crossRef.empty': {
    pt: 'Nada deste tipo no documento.',
    en: 'Nothing of this type in the document.',
  },
  'references.crossRef.show': { pt: 'Inserir referência a', en: 'Insert reference to' },
  'references.crossRef.showText': { pt: 'Texto', en: 'Text' },
  'references.crossRef.showNumber': { pt: 'Número', en: 'Number' },
  'references.crossRef.showPage': { pt: 'Número da página', en: 'Page number' },
  'references.crossRef.link': { pt: 'Inserir como hiperlink', en: 'Insert as hyperlink' },
} satisfies Catalog
