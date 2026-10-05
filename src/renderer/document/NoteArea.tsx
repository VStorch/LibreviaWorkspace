import { useLayoutEffect, useRef } from 'react'
import { noteBody, parkNoteBody, placeNoteBody, subscribeNoteBodies } from './extensions/note-view.js'
import { NOTE_SEPARATOR_PX, type NoteArea, type NoteAreaItem } from './usePagination.js'

/**
 * The first-line slice shows the note's **own** editable body, clipped; the continuation is a
 * shifted copy, because an element cannot live in two places. Clicking the continuation takes the
 * cursor to the body.
 */
export function NoteAreaView({ area }: { area: NoteArea }): React.JSX.Element {
  return (
    <div
      className={`paper-notes paper-notes--${area.kind}`}
      style={{ top: `${area.topPx}px`, left: `${area.leftPx}px`, width: `${area.widthPx}px` }}
    >
      {area.separator !== null && (
        <div
          className={`paper-notes__separator${area.separator === 'continuation' ? ' paper-notes__separator--continued' : ''}`}
          style={{ height: `${NOTE_SEPARATOR_PX}px` }}
          aria-hidden="true"
        />
      )}
      {area.items.map((item) =>
        item.fromLine === 0 ? (
          <LiveNote key={item.key} item={item} />
        ) : (
          <ContinuedNote key={`${item.key}:${item.fromLine}`} item={item} />
        ),
      )}
    </div>
  )
}

function LiveNote({ item }: { item: NoteAreaItem }): React.JSX.Element {
  const slot = useRef<HTMLDivElement>(null)
  // Only when the note moves: redoing it on every draw would steal focus.
  useLayoutEffect(() => {
    const element = slot.current
    const body = noteBody(item.key)
    if (element === null || body === undefined) return undefined
    placeNoteBody(body, element)
    return () => parkNoteBody(body, element)
  }, [item.key])

  return <div ref={slot} className="paper-notes__slot" style={{ height: `${item.heightPx}px` }} />
}

function ContinuedNote({ item }: { item: NoteAreaItem }): React.JSX.Element {
  const slot = useRef<HTMLDivElement>(null)

  // The copy is redone outside React: state subscribed on every key press would make React give up
  // (error 185).
  useLayoutEffect(() => {
    const element = slot.current
    if (element === null) return undefined
    const copy = (): void => {
      const body = noteBody(item.key)
      if (body === undefined) return
      const clone = body.body.cloneNode(true) as HTMLElement
      clone.removeAttribute('contenteditable')
      clone.setAttribute('aria-hidden', 'true')
      clone.style.marginTop = `${-item.clipTopPx}px`
      element.replaceChildren(clone)
    }
    copy()
    let frame = 0
    const unsubscribe = subscribeNoteBodies(() => {
      if (frame === 0)
        frame = requestAnimationFrame(() => {
          frame = 0
          copy()
        })
    })
    return () => {
      unsubscribe()
      if (frame !== 0) cancelAnimationFrame(frame)
    }
  }, [item.key, item.clipTopPx])

  return (
    <div
      ref={slot}
      className="paper-notes__slot paper-notes__slot--continued"
      style={{ height: `${item.heightPx}px` }}
      onMouseDown={(event) => {
        event.preventDefault()
        noteBody(item.key)?.reveal()
      }}
    />
  )
}
