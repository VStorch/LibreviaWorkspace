import type { Catalog } from '../message.js'

export const REVISIONS = {
  'revisions.accept': { pt: 'Aceitar alteração', en: 'Accept change' },
  'revisions.reject': { pt: 'Rejeitar alteração', en: 'Reject change' },
  'revisions.acceptAll': { pt: 'Aceitar todas as alterações', en: 'Accept all changes' },
  'revisions.rejectAll': { pt: 'Rejeitar todas as alterações', en: 'Reject all changes' },
  'revisions.next': { pt: 'Próxima alteração', en: 'Next change' },
  'revisions.previous': { pt: 'Alteração anterior', en: 'Previous change' },
  'revisions.track': { pt: 'Controlar alterações', en: 'Track changes' },
  'revisions.trackHint': {
    pt: 'Marca o texto inserido e excluído com o seu nome. Mudanças de formatação não são controladas.',
    en: 'Marks inserted and deleted text with your name. Formatting changes are not tracked.',
  },
  'revisions.trackOn': { pt: 'Controle de alterações: ativado', en: 'Track changes: on' },
  'revisions.unknownAuthor': { pt: 'Autor', en: 'Author' },
  'revisions.show': { pt: 'Mostrar', en: 'Show' },
  'revisions.show.all': { pt: 'Marcação completa', en: 'All markup' },
  'revisions.show.simple': { pt: 'Marcação simples', en: 'Simple markup' },
  'revisions.show.none': { pt: 'Sem marcação', en: 'No markup' },
  'revisions.show.original': { pt: 'Original', en: 'Original' },
} satisfies Catalog
