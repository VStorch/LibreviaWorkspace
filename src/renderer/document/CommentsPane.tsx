import { useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react'
import type { Editor } from '@tiptap/react'
import { TextSelection } from '@tiptap/pm/state'
import type { DocumentComment } from '@services/document/model.js'
import { useLanguage, useT } from '../i18n.js'
import { useWorkspace } from '../state/workspace.js'
import {
  cancelNewComment,
  deleteCommentThread,
  editComment,
  replyToComment,
  setCommentDone,
} from './comment-commands.js'
import { commentAnchorsOf, focusComment } from './extensions/comment.js'

/** Vão entre dois cartões empilhados, em pixels da folha. */
const GAP_PX = 8

/** A distância entre a borda direita das folhas e a coluna dos cartões. */
export const COMMENTS_PANE_OFFSET_PX = 24

/** A largura que a coluna ocupa ao lado das folhas — ver `.comments-pane` no CSS. */
export const COMMENTS_PANE_WIDTH_PX = 260 + COMMENTS_PANE_OFFSET_PX

interface Thread {
  readonly root: DocumentComment
  readonly replies: readonly DocumentComment[]
}

/** As conversas: cada comentário que abre uma, com as respostas na ordem do arquivo. */
function threadsOf(comments: readonly DocumentComment[]): Thread[] {
  const known = new Set(comments.map((comment) => comment.id))
  const replies = new Map<string, DocumentComment[]>()
  const roots: DocumentComment[] = []
  for (const comment of comments) {
    // A resposta cujo comentário não está no arquivo vira conversa própria: some
    // da tela é que ela não pode.
    if (comment.parentId !== undefined && known.has(comment.parentId)) {
      replies.set(comment.parentId, [...(replies.get(comment.parentId) ?? []), comment])
    } else {
      roots.push(comment)
    }
  }
  return roots.map((root) => ({ root, replies: replies.get(root.id) ?? [] }))
}

/** A caixa de texto aberta num cartão: a resposta, ou o texto do comentário. */
interface Composing {
  readonly kind: 'reply' | 'edit'
  readonly id: string
}

/**
 * Os comentários do documento, numa coluna ao lado das folhas (M10).
 *
 * `comments` são os que valem agora (`resolveComments`): o desfeito e o excluído
 * já não vêm. O cartão escolhido mostra as ações — responder, editar, resolver,
 * excluir —, e o recém-inserido (`commentDraft`) abre com a caixa de texto.
 *
 * Cada cartão fica na altura do trecho que comenta — a da ponta de início,
 * ou a do fim no comentário de ponto — e desce o quanto for preciso para não
 * cobrir o de cima, como a margem do Word. Clicar seleciona o trecho e o realça.
 *
 * A medida é a mesma da paginação: uma por quadro, no máximo, disparada pela
 * edição e pela mudança de tamanho da pilha. O que o arquivo ancora fora do
 * corpo (cabeçalho, nota, caixa de texto) vai para o fim da coluna.
 */
export function CommentsPane({
  editor,
  comments,
  outside,
  leftPx,
}: {
  readonly editor: Editor
  readonly comments: readonly DocumentComment[]
  /** As conversas ancoradas fora do corpo — ver `commentsOutsideOf`. */
  readonly outside: ReadonlySet<string>
  /** A borda direita das folhas, em pixels da pilha. */
  readonly leftPx: number
}): React.JSX.Element {
  const t = useT()
  const language = useLanguage()
  const readOnly = useWorkspace((state) => state.readOnly)
  const draft = useWorkspace((state) => state.commentDraft)
  const paneRef = useRef<HTMLElement>(null)
  const listRef = useRef<HTMLOListElement>(null)
  const threads = useMemo(() => threadsOf(comments), [comments])
  const [tops, setTops] = useState<ReadonlyMap<string, number>>(new Map())
  const [active, setActive] = useState<string | null>(null)
  const [composing, setComposing] = useState<Composing | null>(null)

  // O comentário recém-inserido abre escolhido, com a caixa de texto.
  useEffect(() => {
    if (draft === null) return
    setActive(draft)
    setComposing(null)
    editor.view.dispatch(focusComment(editor.state.tr, { active: draft }))
  }, [editor, draft])

  // O rascunho cujas pontas o desfazer levou não espera mais texto.
  useEffect(() => {
    if (draft !== null && !threads.some((thread) => thread.root.id === draft)) {
      useWorkspace.getState().setCommentDraft(null)
    }
  }, [draft, threads])

  // As conversas resolvidas saem do realce do texto, como no Word.
  useEffect(() => {
    const resolved = new Set(threads.filter((thread) => thread.root.done).map((thread) => thread.root.id))
    editor.view.dispatch(focusComment(editor.state.tr, { resolved }))
  }, [editor, threads])

  useEffect(() => {
    const host = paneRef.current?.parentElement
    if (host === null || host === undefined) return

    const measure = (): void => {
      if (editor.isDestroyed) return
      const box = host.getBoundingClientRect()
      // O zoom é `transform` na pilha: a tela mede escalado, a pilha não.
      const scale = host.offsetWidth > 0 ? box.width / host.offsetWidth : 1
      const anchors = commentAnchorsOf(editor.state.doc)
      const next = new Map<string, number>()
      for (const { root } of threads) {
        const anchor = anchors.get(root.id)
        const pos = anchor?.start ?? anchor?.end ?? null
        if (pos === null) continue
        try {
          next.set(root.id, (editor.view.coordsAtPos(pos).top - box.top) / scale)
        } catch {
          // A posição saiu do documento entre a edição e o quadro: a medida
          // seguinte a acha.
        }
      }
      setTops((previous) => (sameTops(previous, next) ? previous : next))
    }

    let scheduled = 0
    const schedule = (): void => {
      if (scheduled !== 0) return
      scheduled = requestAnimationFrame(() => {
        scheduled = 0
        measure()
      })
    }

    schedule()
    // A transação, e não só a edição: os vãos da paginação e as decorações
    // descem o texto sem mudar o documento.
    editor.on('transaction', schedule)
    const observer = new ResizeObserver(schedule)
    observer.observe(host)
    return () => {
      editor.off('transaction', schedule)
      observer.disconnect()
      if (scheduled !== 0) cancelAnimationFrame(scheduled)
    }
  }, [editor, threads])

  // Empilha: cada cartão na altura do trecho, ou logo abaixo do de cima. Depois
  // de todo desenho, porque a altura de um cartão só existe depois dele.
  useLayoutEffect(() => {
    let bottom = 0
    for (const card of Array.from(listRef.current?.children ?? [])) {
      if (!(card instanceof HTMLElement)) continue
      const wanted = Number(card.dataset['y'])
      const top = Math.max(Number.isFinite(wanted) ? wanted : bottom, bottom)
      card.style.top = `${top}px`
      bottom = top + card.offsetHeight + GAP_PX
    }
  })

  const ordered = useMemo(() => {
    const anchored = threads.filter((thread) => tops.has(thread.root.id))
    anchored.sort((left, right) => tops.get(left.root.id)! - tops.get(right.root.id)!)
    return [...anchored, ...threads.filter((thread) => !tops.has(thread.root.id))]
  }, [threads, tops])

  function choose(cid: string): void {
    const next = active === cid ? null : cid
    setActive(next)
    let tr = focusComment(editor.state.tr, { active: next })
    const anchor = commentAnchorsOf(editor.state.doc).get(cid)
    if (next !== null && anchor !== undefined) {
      // Depois da ponta de início e antes da de fim: o trecho, e só ele.
      const from = anchor.start === null ? (anchor.end ?? 0) : anchor.start + 1
      const to = anchor.end ?? from
      tr = tr.setSelection(TextSelection.create(tr.doc, from, Math.max(from, to))).scrollIntoView()
    }
    editor.view.dispatch(tr)
  }

  const date = (value: string): string => {
    const parsed = value === '' ? Number.NaN : Date.parse(value)
    return Number.isNaN(parsed)
      ? ''
      : new Date(parsed).toLocaleString(language, { dateStyle: 'short', timeStyle: 'short' })
  }

  const body = (comment: DocumentComment): React.JSX.Element => (
    <>
      <div className="comment-card__head">
        <span className="comment-card__author">{comment.author || t('comments.card.unknownAuthor')}</span>
        <span className="comment-card__date">{date(comment.date)}</span>
      </div>
      {composing?.kind === 'edit' && composing.id === comment.id ? (
        <CommentComposer
          initial={comment.paragraphs.join('\n')}
          submitLabel={t('comments.action.save')}
          onSubmit={(text) => {
            setComposing(null)
            editComment(comment.id, text)
          }}
          onCancel={() => setComposing(null)}
        />
      ) : (
        comment.paragraphs.map((text, index) => (
          <p key={index} className="comment-card__text">
            {text}
          </p>
        ))
      )}
      {comment.rich === true && <p className="comment-card__note">{t('comments.card.rich')}</p>}
    </>
  )

  /** O botão do cartão não escolhe o cartão: o clique para nele. */
  const action = (label: string, run: () => void): React.JSX.Element => (
    <button
      type="button"
      className="comment-card__action"
      onClick={(event) => {
        event.stopPropagation()
        run()
      }}
    >
      {label}
    </button>
  )

  return (
    <aside
      ref={paneRef}
      className="comments-pane"
      style={{ left: `${leftPx + COMMENTS_PANE_OFFSET_PX}px` }}
      aria-label={t('comments.pane.title')}
    >
      <ol ref={listRef} className="comments-pane__list">
        {ordered.map(({ root, replies }) => {
          const isActive = active === root.id
          const isDraft = draft === root.id
          // A resolvida fica recolhida — só o cabeçalho — até ser escolhida.
          const collapsed = root.done && !isActive
          return (
            <li
              key={root.id}
              data-y={tops.get(root.id) ?? ''}
              data-cid={root.id}
              className={`comment-card${isActive ? ' comment-card--active' : ''}${root.done ? ' comment-card--done' : ''}`}
              role="button"
              tabIndex={0}
              aria-pressed={isActive}
              aria-label={t('comments.card.label', {
                author: root.author || t('comments.card.unknownAuthor'),
              })}
              onClick={() => {
                if (!isDraft) choose(root.id)
              }}
              onKeyDown={(event) => {
                // A tecla da caixa de texto e dos botões é deles.
                if (event.target !== event.currentTarget) return
                if (event.key !== 'Enter' && event.key !== ' ') return
                event.preventDefault()
                choose(root.id)
              }}
            >
              {collapsed ? (
                <div className="comment-card__head">
                  <span className="comment-card__author">
                    {root.author || t('comments.card.unknownAuthor')}
                  </span>
                  <span className="comment-card__badge">{t('comments.card.resolved')}</span>
                </div>
              ) : isDraft ? (
                <>
                  <div className="comment-card__head">
                    <span className="comment-card__author">
                      {root.author || t('comments.card.unknownAuthor')}
                    </span>
                  </div>
                  <CommentComposer
                    initial=""
                    submitLabel={t('comments.action.post')}
                    placeholder={t('comments.editor.placeholder')}
                    onSubmit={(text) => {
                      if (text.trim() === '') cancelNewComment(editor, root.id)
                      else editComment(root.id, text)
                    }}
                    onCancel={() => cancelNewComment(editor, root.id)}
                  />
                </>
              ) : (
                <>
                  {body(root)}
                  {root.done && <p className="comment-card__note">{t('comments.card.resolved')}</p>}
                  {outside.has(root.id) && (
                    <p className="comment-card__note">{t('comments.card.unanchored')}</p>
                  )}
                  {replies.length > 0 && (
                    <ol className="comment-card__replies" aria-label={t('comments.card.replies')}>
                      {replies.map((reply) => (
                        <li key={reply.id}>{body(reply)}</li>
                      ))}
                    </ol>
                  )}
                  {composing?.kind === 'reply' && composing.id === root.id && (
                    <CommentComposer
                      initial=""
                      submitLabel={t('comments.action.reply')}
                      placeholder={t('comments.reply.placeholder')}
                      onSubmit={(text) => {
                        setComposing(null)
                        if (text.trim() !== '') replyToComment(root.id, text)
                      }}
                      onCancel={() => setComposing(null)}
                    />
                  )}
                </>
              )}
              {isActive && !isDraft && !readOnly && composing === null && (
                <div className="comment-card__actions">
                  {!root.done &&
                    action(t('comments.action.reply'), () => setComposing({ kind: 'reply', id: root.id }))}
                  {!root.done &&
                    action(t('comments.action.edit'), () => setComposing({ kind: 'edit', id: root.id }))}
                  {action(t(root.done ? 'comments.action.reopen' : 'comments.action.resolve'), () =>
                    setCommentDone(root.id, !root.done),
                  )}
                  {action(t('comments.action.delete'), () => {
                    setActive(null)
                    deleteCommentThread(editor, root.id)
                  })}
                </div>
              )}
            </li>
          )
        })}
      </ol>
    </aside>
  )
}

/**
 * A caixa de texto do cartão: o comentário novo, a resposta, a edição.
 *
 * Abre com o foco. `Ctrl+Enter` confirma e `Esc` desiste; o clique e a tecla não
 * sobem para o cartão, que os leria como "escolher".
 */
function CommentComposer({
  initial,
  submitLabel,
  placeholder,
  onSubmit,
  onCancel,
}: {
  readonly initial: string
  readonly submitLabel: string
  readonly placeholder?: string
  readonly onSubmit: (text: string) => void
  readonly onCancel: () => void
}): React.JSX.Element {
  const t = useT()
  const [text, setText] = useState(initial)
  const ref = useRef<HTMLTextAreaElement>(null)

  useEffect(() => {
    const field = ref.current
    if (field === null) return
    field.focus()
    field.setSelectionRange(field.value.length, field.value.length)
  }, [])

  return (
    <div
      className="comment-composer"
      onClick={(event) => event.stopPropagation()}
      onKeyDown={(event) => event.stopPropagation()}
    >
      <textarea
        ref={ref}
        className="comment-composer__field"
        aria-label={t('comments.editor.label')}
        placeholder={placeholder}
        value={text}
        rows={3}
        onChange={(event) => setText(event.target.value)}
        onKeyDown={(event) => {
          if (event.key === 'Escape') {
            event.preventDefault()
            onCancel()
          } else if (event.key === 'Enter' && (event.ctrlKey || event.metaKey)) {
            event.preventDefault()
            onSubmit(text)
          }
        }}
      />
      <div className="comment-composer__buttons">
        <button type="button" className="btn" onClick={onCancel}>
          {t('comments.action.cancel')}
        </button>
        <button type="button" className="btn btn--primary" onClick={() => onSubmit(text)}>
          {submitLabel}
        </button>
      </div>
    </div>
  )
}

function sameTops(left: ReadonlyMap<string, number>, right: ReadonlyMap<string, number>): boolean {
  if (left.size !== right.size) return false
  for (const [cid, top] of left) {
    // Meio pixel não vale um redesenho: o arredondamento oscila sozinho.
    if (!right.has(cid) || Math.abs(right.get(cid)! - top) > 0.5) return false
  }
  return true
}
