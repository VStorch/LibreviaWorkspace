import type { Catalog } from '../message.js'

/**
 * Comentários (M10): o painel ao lado da folha.
 *
 * Área própria porque o marco inteiro mora aqui — e as fases seguintes (criar,
 * responder, resolver) crescem a lista sem tocar as outras áreas.
 */
export const COMMENTS = {
  'comments.pane.title': { pt: 'Comentários', en: 'Comments' },
  'comments.card.label': { pt: 'Comentário de {author}', en: 'Comment by {author}' },
  'comments.card.unknownAuthor': { pt: 'Autor desconhecido', en: 'Unknown author' },
  'comments.card.replies': { pt: 'Respostas', en: 'Replies' },
  'comments.card.resolved': { pt: 'Resolvido', en: 'Resolved' },
  'comments.card.rich': {
    pt: 'Este comentário tem formatação ou imagem que o painel não mostra. Ela continua no arquivo.',
    en: 'This comment has formatting or images the pane does not show. They stay in the file.',
  },
  'comments.card.unanchored': {
    pt: 'Sem trecho no corpo do texto: foi apagado, ou está num cabeçalho, nota ou caixa de texto.',
    en: 'No range in the body text: it was deleted, or it is in a header, note or text box.',
  },
} satisfies Catalog
