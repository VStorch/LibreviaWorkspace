import type { Editor } from '@tiptap/react'
import type { DocumentComment } from '@services/document/model.js'
import { initialsOf, nextCommentId, paragraphsOfText, resolveComments } from '@services/document/comments.js'
import { t } from '../i18n.js'
import { currentPreferences, setPreference } from '../state/preferences.js'
import { useWorkspace } from '../state/workspace.js'
import {
  adjacentComment,
  commentAnchorsOf,
  commentSelectionOf,
  commentsKey,
  insertCommentAnchors,
  removeCommentAnchors,
  selectComment,
} from './extensions/comment.js'
import { caretOf, selectInNote } from './extensions/note-view.js'

/**
 * Criar, responder, editar, resolver e excluir comentários.
 *
 * O desenho é o das seções (ver section-commands.ts): o comentário mora em dois
 * lugares — as pontas no texto e o corpo na biblioteca da loja —, e quem diz qual
 * vale é o texto (`resolveComments`). Inserir e excluir são **uma transação do
 * editor** cada, e o desfazer tira e devolve o cartão junto com as pontas. Texto,
 * resposta e resolvido mudam só a loja: ficam fora do desfazer, como os estilos,
 * e marcam o documento.
 */

/**
 * Comentários só se criam e mudam em documento que os grava.
 *
 * O rascunho de antes deles (`.sdoc` < 7) não tem as pontas nos nós, e a gravação
 * dele deixa as partes de comentário como estão: a mudança apareceria na tela e
 * sumiria no arquivo. O comando recusa e diz por quê.
 */
export function commentEditsAllowed(): boolean {
  const store = useWorkspace.getState()
  if (store.readOnly) return false
  if (!store.beforeComments) return true
  store.showError({ code: 'INTERNAL', message: t('comments.legacyDraft') })
  return false
}

/** Um comentário novo, assinado pelo autor das preferências e datado de agora. */
function newComment(library: readonly DocumentComment[], parentId?: string): DocumentComment {
  const author = currentPreferences().authorName.trim()
  const initials = initialsOf(author)
  return {
    id: nextCommentId(library),
    author,
    ...(initials === '' ? {} : { initials }),
    date: new Date().toISOString(),
    paragraphs: [],
    done: false,
    ...(parentId === undefined ? {} : { parentId }),
  }
}

/**
 * Inserir → Comentário: as pontas na seleção e um cartão novo no painel, com a
 * caixa de texto aberta. Vazio, o cartão desiste e leva as pontas (`cancelNewComment`).
 */
export function insertComment(editor: Editor): void {
  if (!commentEditsAllowed()) return
  const store = useWorkspace.getState()
  const wasClean = !store.isDirty
  const comment = newComment(store.comments)
  // A biblioteca antes do texto, como nas seções: a ponta nova precisa achar o
  // corpo quando o painel medir. Entrada a mais não muda nada até o texto apontá-la.
  // Sem `focus()`: com o editor fora de foco (o clique no menu de contexto), o
  // foco do TipTap chega num quadro seguinte e rouba o da caixa do cartão.
  // Digitando numa nota, as pontas vão para o corpo dela, na seleção dele.
  const caret = caretOf(editor.view)
  // Comentário novo dentro de nota, não: o LibreOffice não abre o .docx que traz
  // `w:commentReference` em `footnotes.xml` ("não foi possível carregar"), e o
  // próprio LibreOffice os descarta ao gravar. O que vem do arquivo continua lido,
  // mostrado e devolvido como estava.
  if (caret.note !== null) {
    store.showError({ code: 'INTERNAL', message: t('comments.notInNote') })
    return
  }
  const inserted = editor
    .chain()
    .command(({ tr, state }) => {
      if (!insertCommentAnchors(tr, state.schema, comment.id, caret)) return false
      store.setComments([...store.comments, comment])
      return true
    })
    .run()
  if (!inserted) return
  showCommentsPane()
  cleanBeforeDraft = wasClean
  store.setCommentDraft(comment.id)
}

/** O documento estava salvo quando o cartão novo abriu: desistir dele o devolve assim. */
let cleanBeforeDraft = false

/**
 * O cartão novo desistiu (Esc, ou confirmado vazio): as pontas saem, fora do
 * desfazer — refazer um comentário que nunca teve texto seria um cartão vazio.
 */
export function cancelNewComment(editor: Editor, cid: string): void {
  const store = useWorkspace.getState()
  store.setCommentDraft(null)
  const tr = editor.state.tr
  if (removeCommentAnchors(tr, cid)) editor.view.dispatch(tr.setMeta('addToHistory', false))
  if (cleanBeforeDraft) useWorkspace.setState({ isDirty: false })
  cleanBeforeDraft = false
}

/** Excluir a conversa: as pontas saem do texto, e com elas o cartão e as respostas. */
export function deleteCommentThread(editor: Editor, cid: string): void {
  if (!commentEditsAllowed()) return
  const tr = editor.state.tr
  if (removeCommentAnchors(tr, cid)) editor.view.dispatch(tr)
  const store = useWorkspace.getState()
  if (store.commentDraft === cid) store.setCommentDraft(null)
}

/** Troca o texto de um comentário. */
export function editComment(id: string, text: string): void {
  if (!commentEditsAllowed()) return
  const store = useWorkspace.getState()
  const paragraphs = paragraphsOfText(text)
  store.setComments(
    store.comments.map((comment) => (comment.id === id ? { ...comment, paragraphs } : comment)),
  )
  if (store.commentDraft === id) store.setCommentDraft(null)
}

/** Uma resposta nova no fim da conversa `rootId`. */
export function replyToComment(rootId: string, text: string): void {
  if (!commentEditsAllowed()) return
  const store = useWorkspace.getState()
  const reply = { ...newComment(store.comments, rootId), paragraphs: paragraphsOfText(text) }
  store.setComments([...store.comments, reply])
}

/** Resolver e reabrir: vale para a conversa inteira, pelo comentário que a abre. */
export function setCommentDone(rootId: string, done: boolean): void {
  if (!commentEditsAllowed()) return
  const store = useWorkspace.getState()
  store.setComments(store.comments.map((comment) => (comment.id === rootId ? { ...comment, done } : comment)))
}

/** O painel escondido volta quando um comando de comentário precisa dele. */
function showCommentsPane(): void {
  if (!currentPreferences().commentsPane) void setPreference({ commentsPane: true })
}

/**
 * Próximo e anterior: o cursor vai ao trecho da conversa seguinte (ou da
 * anterior) pela ordem do texto, e o cartão dela fica escolhido — ver
 * `adjacentComment`. Devolve se achou alguma.
 */
export function goToComment(editor: Editor, direction: 1 | -1): boolean {
  const store = useWorkspace.getState()
  const anchors = commentAnchorsOf(editor.state.doc)
  const visible = resolveComments(new Set(anchors.keys()), store.comments, new Set(store.commentsOutside))
  const ids = new Set(visible.map((comment) => comment.id))
  const threads = new Set(
    visible
      .filter((comment) => comment.parentId === undefined || !ids.has(comment.parentId))
      .map((comment) => comment.id),
  )
  const active = commentsKey.getState(editor.state)?.active ?? null
  const cid = adjacentComment(editor.state.doc, threads, active, caretOf(editor.view).from, direction)
  if (cid === null) return false
  showCommentsPane()
  showComment(editor, cid, true)
  return true
}

/**
 * Escolhe a conversa `cid`: o realce em foco e o trecho selecionado — no corpo da
 * nota, quando é lá que ela está (`selectInNote`). `focus` leva o teclado ao
 * trecho; o clique no cartão deixa o foco onde está.
 */
export function showComment(editor: Editor, cid: string, focus: boolean): void {
  const anchor = commentAnchorsOf(editor.state.doc).get(cid)
  const range = anchor === undefined ? null : commentSelectionOf(anchor)
  if (focus) editor.view.focus()
  editor.view.dispatch(selectComment(editor.state.tr, cid))
  if (range !== null) selectInNote(editor.view, range.from, range.to, focus)
}
