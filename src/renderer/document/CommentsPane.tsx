import { useEffect, useLayoutEffect, useMemo, useRef, useState, type RefObject } from 'react'
import { useEditorState, type Editor } from '@tiptap/react'
import type { DocumentComment } from '@services/document/model.js'
import { useLanguage, useT } from '../i18n.js'
import { useWorkspace } from '../state/workspace.js'
import {
  cancelNewComment,
  deleteCommentThread,
  editComment,
  replyToComment,
  setCommentDone,
  showComment,
} from './comment-commands.js'
import { commentAnchorsOf, commentsKey, focusComment } from './extensions/comment.js'
import { coordsInDocument } from './extensions/note-view.js'

/** Sheet pixels. */
const GAP_PX = 8

export const COMMENTS_PANE_OFFSET_PX = 24

/** See `.comments-pane` in the CSS. */
export const COMMENTS_PANE_WIDTH_PX = 260 + COMMENTS_PANE_OFFSET_PX

interface Thread {
  readonly root: DocumentComment
  readonly replies: readonly DocumentComment[]
}

/** With replies in file order. */
function threadsOf(comments: readonly DocumentComment[]): Thread[] {
  const known = new Set(comments.map((comment) => comment.id))
  const replies = new Map<string, DocumentComment[]>()
  const roots: DocumentComment[] = []
  for (const comment of comments) {
    // A reply whose comment is not in the file becomes its own thread: it must not vanish from the
    // screen.
    if (comment.parentId !== undefined && known.has(comment.parentId)) {
      replies.set(comment.parentId, [...(replies.get(comment.parentId) ?? []), comment])
    } else {
      roots.push(comment)
    }
  }
  return roots.map((root) => ({ root, replies: replies.get(root.id) ?? [] }))
}

interface Composing {
  readonly kind: 'reply' | 'edit'
  readonly id: string
}

/**
 * `comments` are the ones that count now (`resolveComments`). Each card sits at the height of its
 * range and moves down just enough not to cover the one above, like Word's margin. One measurement
 * per frame, like pagination; threads the file anchors outside the body go to the end.
 */
export function CommentsPane({
  editor,
  comments,
  outside,
  leftPx,
}: {
  readonly editor: Editor
  readonly comments: readonly DocumentComment[]
  readonly outside: ReadonlySet<string>
  /** Stack pixels. */
  readonly leftPx: number
}): React.JSX.Element {
  const t = useT()
  const readOnly = useWorkspace((state) => state.readOnly)
  const draft = useWorkspace((state) => state.commentDraft)
  const paneRef = useRef<HTMLElement>(null)
  const listRef = useRef<HTMLOListElement>(null)
  const threads = useMemo(() => threadsOf(comments), [comments])
  // The focused thread lives in the highlight: the menu's Next picks it without going through the
  // pane.
  const active = useEditorState({
    editor,
    selector: ({ editor: current }) => commentsKey.getState(current.state)?.active ?? null,
  })
  const [composing, setComposing] = useState<Composing | null>(null)
  useThreadState(editor, draft, threads, setComposing)
  const tops = useCommentTops(editor, threads, paneRef)
  useStackedCards(listRef)

  const ordered = useMemo(() => {
    const anchored = threads.filter((thread) => tops.has(thread.root.id))
    anchored.sort((left, right) => tops.get(left.root.id)! - tops.get(right.root.id)!)
    return [...anchored, ...threads.filter((thread) => !tops.has(thread.root.id))]
  }, [threads, tops])

  const state: CardState = {
    editor,
    active,
    draft,
    readOnly,
    composing,
    setComposing,
    outside,
    choose: (cid) => {
      setComposing(null)
      if (active === cid) editor.view.dispatch(focusComment(editor.state.tr, { active: null }))
      else showComment(editor, cid, false)
    },
    setActive: (cid) => editor.view.dispatch(focusComment(editor.state.tr, { active: cid })),
  }

  return (
    <aside
      ref={paneRef}
      className="comments-pane"
      style={{ left: `${leftPx + COMMENTS_PANE_OFFSET_PX}px` }}
      aria-label={t('comments.pane.title')}
    >
      <ol ref={listRef} className="comments-pane__list">
        {ordered.map((thread) => (
          <CommentCard key={thread.root.id} thread={thread} top={tops.get(thread.root.id)} state={state} />
        ))}
      </ol>
    </aside>
  )
}

/** The draft, the focus and resolved threads, carried to the editor highlight. */
function useThreadState(
  editor: Editor,
  draft: string | null,
  threads: readonly Thread[],
  setComposing: (composing: Composing | null) => void,
): void {
  useEffect(() => {
    if (draft === null) return
    setComposing(null)
    editor.view.dispatch(focusComment(editor.state.tr, { active: draft }))
  }, [editor, draft, setComposing])

  // A draft whose ends undo took away no longer waits for text.
  useEffect(() => {
    if (draft !== null && !threads.some((thread) => thread.root.id === draft)) {
      useWorkspace.getState().setCommentDraft(null)
    }
  }, [draft, threads])

  useEffect(() => {
    const resolved = new Set(threads.filter((thread) => thread.root.done).map((thread) => thread.root.id))
    editor.view.dispatch(focusComment(editor.state.tr, { resolved }))
  }, [editor, threads])
}

/** In pane pixels. */
function useCommentTops(
  editor: Editor,
  threads: readonly Thread[],
  paneRef: RefObject<HTMLElement | null>,
): ReadonlyMap<string, number> {
  const [tops, setTops] = useState<ReadonlyMap<string, number>>(new Map())
  useEffect(() => {
    const host = paneRef.current?.parentElement
    if (host === null || host === undefined) return

    const measure = (): void => {
      if (editor.isDestroyed) return
      const box = host.getBoundingClientRect()
      // Zoom is a `transform`: the screen measures scaled.
      const scale = host.offsetWidth > 0 ? box.width / host.offsetWidth : 1
      const anchors = commentAnchorsOf(editor.state.doc)
      const next = new Map<string, number>()
      for (const { root } of threads) {
        const anchor = anchors.get(root.id)
        const pos = anchor?.start ?? anchor?.end ?? null
        if (pos === null) continue
        try {
          next.set(root.id, (coordsInDocument(editor.view, pos).top - box.top) / scale)
        } catch {
          // The position left the document between the edit and the frame: the next measurement
          // finds it.
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
    // The transaction, not only the edit: gaps and decorations move the text down.
    editor.on('transaction', schedule)
    const observer = new ResizeObserver(schedule)
    observer.observe(host)
    return () => {
      editor.off('transaction', schedule)
      observer.disconnect()
      if (scheduled !== 0) cancelAnimationFrame(scheduled)
    }
  }, [editor, threads, paneRef])
  return tops
}

/** After drawing: a card's height only exists afterwards. */
function useStackedCards(listRef: RefObject<HTMLOListElement | null>): void {
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
}

interface CardState {
  readonly editor: Editor
  readonly active: string | null
  readonly draft: string | null
  readonly readOnly: boolean
  readonly composing: Composing | null
  readonly setComposing: (composing: Composing | null) => void
  readonly outside: ReadonlySet<string>
  readonly choose: (cid: string) => void
  readonly setActive: (cid: string | null) => void
}

function CommentCard({
  thread,
  top,
  state,
}: {
  thread: Thread
  top: number | undefined
  state: CardState
}): React.JSX.Element {
  const t = useT()
  const { root } = thread
  const isActive = state.active === root.id
  const isDraft = state.draft === root.id
  const collapsed = root.done && !isActive
  return (
    <li
      data-y={top ?? ''}
      data-cid={root.id}
      className={`comment-card${isActive ? ' comment-card--active' : ''}${root.done ? ' comment-card--done' : ''}`}
      role="button"
      tabIndex={0}
      aria-pressed={isActive}
      aria-label={t('comments.card.label', { author: root.author || t('comments.card.unknownAuthor') })}
      onClick={() => {
        if (!isDraft) state.choose(root.id)
      }}
      onKeyDown={(event) => {
        if (event.target !== event.currentTarget) return
        if (event.key !== 'Enter' && event.key !== ' ') return
        event.preventDefault()
        state.choose(root.id)
      }}
    >
      {collapsed ? (
        <div className="comment-card__head">
          <span className="comment-card__author">{root.author || t('comments.card.unknownAuthor')}</span>
          <span className="comment-card__badge">{t('comments.card.resolved')}</span>
        </div>
      ) : isDraft ? (
        <DraftCard root={root} editor={state.editor} />
      ) : (
        <OpenThread thread={thread} state={state} />
      )}
      {isActive && !isDraft && !state.readOnly && state.composing === null && (
        <ThreadActions root={root} state={state} />
      )}
    </li>
  )
}

function DraftCard({ root, editor }: { root: DocumentComment; editor: Editor }): React.JSX.Element {
  const t = useT()
  return (
    <>
      <div className="comment-card__head">
        <span className="comment-card__author">{root.author || t('comments.card.unknownAuthor')}</span>
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
  )
}

function OpenThread({ thread, state }: { thread: Thread; state: CardState }): React.JSX.Element {
  const t = useT()
  const { root, replies } = thread
  const { composing, setComposing } = state
  const canEditReplies = state.active === root.id && !root.done && !state.readOnly && composing === null
  return (
    <>
      <CommentBody comment={root} state={state} />
      {root.done && <p className="comment-card__note">{t('comments.card.resolved')}</p>}
      {state.outside.has(root.id) && <p className="comment-card__note">{t('comments.card.unanchored')}</p>}
      {replies.length > 0 && (
        <ol className="comment-card__replies" aria-label={t('comments.card.replies')}>
          {replies.map((reply) => (
            <li key={reply.id}>
              <CommentBody comment={reply} state={state} />
              {canEditReplies && (
                <div className="comment-card__actions">
                  <CardAction
                    label={t('comments.action.edit')}
                    run={() => setComposing({ kind: 'edit', id: reply.id })}
                  />
                </div>
              )}
            </li>
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
  )
}

function ThreadActions({ root, state }: { root: DocumentComment; state: CardState }): React.JSX.Element {
  const t = useT()
  return (
    <div className="comment-card__actions">
      {!root.done && (
        <CardAction
          label={t('comments.action.reply')}
          run={() => state.setComposing({ kind: 'reply', id: root.id })}
        />
      )}
      {!root.done && (
        <CardAction
          label={t('comments.action.edit')}
          run={() => state.setComposing({ kind: 'edit', id: root.id })}
        />
      )}
      <CardAction
        label={t(root.done ? 'comments.action.reopen' : 'comments.action.resolve')}
        run={() => setCommentDone(root.id, !root.done)}
      />
      <CardAction
        label={t('comments.action.delete')}
        run={() => {
          state.setActive(null)
          deleteCommentThread(state.editor, root.id)
        }}
      />
    </div>
  )
}

function CommentBody({ comment, state }: { comment: DocumentComment; state: CardState }): React.JSX.Element {
  const t = useT()
  const language = useLanguage()
  const { composing, setComposing } = state
  const parsed = comment.date === '' ? Number.NaN : Date.parse(comment.date)
  const date = Number.isNaN(parsed)
    ? ''
    : new Date(parsed).toLocaleString(language, { dateStyle: 'short', timeStyle: 'short' })
  return (
    <>
      <div className="comment-card__head">
        <span className="comment-card__author">{comment.author || t('comments.card.unknownAuthor')}</span>
        <span className="comment-card__date">{date}</span>
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
}

/** Clicking the button does not select the card. */
function CardAction({ label, run }: { label: string; run: () => void }): React.JSX.Element {
  return (
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
}

/** `Ctrl+Enter` confirms and `Esc` gives up; the click and the key do not bubble to the card. */
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
    // Half a pixel is not worth a redraw.
    if (!right.has(cid) || Math.abs(right.get(cid)! - top) > 0.5) return false
  }
  return true
}
