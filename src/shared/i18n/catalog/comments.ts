import type { Catalog } from '../message.js'

/**
 * Comentários (M10): o painel ao lado da folha, e criar, responder, resolver.
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
    pt: 'Ancorado fora do corpo do texto: num cabeçalho, nota ou caixa de texto.',
    en: 'Anchored outside the body text: in a header, note or text box.',
  },
  'comments.new': { pt: 'Novo comentário', en: 'New comment' },
  'comments.editor.label': { pt: 'Texto do comentário', en: 'Comment text' },
  'comments.editor.placeholder': { pt: 'Escreva um comentário…', en: 'Write a comment…' },
  'comments.reply.placeholder': { pt: 'Responder…', en: 'Reply…' },
  'comments.action.post': { pt: 'Comentar', en: 'Comment' },
  'comments.action.save': { pt: 'Salvar', en: 'Save' },
  'comments.action.cancel': { pt: 'Cancelar', en: 'Cancel' },
  'comments.action.reply': { pt: 'Responder', en: 'Reply' },
  'comments.action.edit': { pt: 'Editar', en: 'Edit' },
  'comments.action.delete': { pt: 'Excluir', en: 'Delete' },
  'comments.action.resolve': { pt: 'Resolver', en: 'Resolve' },
  'comments.action.reopen': { pt: 'Reabrir', en: 'Reopen' },
  'comments.author.title': { pt: 'Nome do autor', en: 'Author name' },
  'comments.author.label': { pt: 'Nome', en: 'Name' },
  'comments.author.hint': {
    pt: 'Assina os comentários e as respostas novos. Os que já existem guardam o autor que tinham.',
    en: 'Signs new comments and replies. Existing ones keep the author they had.',
  },
  'comments.legacyDraft': {
    pt: 'Este rascunho é de uma versão anterior aos comentários. Salve-o como .docx e reabra para comentar.',
    en: 'This draft is from a version before comments. Save it as .docx and reopen it to comment.',
  },
} satisfies Catalog
