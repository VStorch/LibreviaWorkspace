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
} satisfies Catalog
