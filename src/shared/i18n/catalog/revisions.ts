import type { Catalog } from '../message.js'

/**
 * Controle de alterações (M10): o menu Revisão e o que o botão direito oferece
 * sobre uma alteração.
 *
 * Área própria pelo mesmo motivo da dos comentários: a fase seguinte (controlar
 * o que se digita) cresce a lista sem tocar as outras áreas.
 */
export const REVISIONS = {
  'revisions.accept': { pt: 'Aceitar alteração', en: 'Accept change' },
  'revisions.reject': { pt: 'Rejeitar alteração', en: 'Reject change' },
  'revisions.acceptAll': { pt: 'Aceitar todas as alterações', en: 'Accept all changes' },
  'revisions.rejectAll': { pt: 'Rejeitar todas as alterações', en: 'Reject all changes' },
  'revisions.next': { pt: 'Próxima alteração', en: 'Next change' },
  'revisions.previous': { pt: 'Alteração anterior', en: 'Previous change' },
} satisfies Catalog
