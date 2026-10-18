import { useEffect } from 'react'
import type { Range } from '@services/spreadsheet/edit.js'
import { fillRange, fillTarget } from '@services/spreadsheet/fill.js'
import type { Sheet } from '@services/spreadsheet/model.js'
import { gridPositionOf } from './grid-position.js'
import { PRIMARY_BUTTON, followSelection, paint, type SelectionConnector } from './useCrossPaneRange.js'

const FILL_HANDLE = '.autofill-handle'

interface Fill {
  readonly source: Range
  target: Range
}

function onFillHandle(event: MouseEvent): boolean {
  const deepest = event.composedPath()[0]
  return (
    event.button === PRIMARY_BUTTON && deepest instanceof Element && deepest.closest(FILL_HANDLE) !== null
  )
}

/**
 * The grid fills only inside the pane where the drag starts and stops at the frozen columns. With
 * frozen columns, the drag of the fill handle is done here, from the whole selection, to wherever
 * the pointer goes.
 */
export function useCrossPaneFill(
  container: { readonly current: HTMLDivElement | null },
  current: { readonly current: Sheet },
  selected: { readonly current: Range },
  applyChange: (sheet: Sheet) => void,
  setRange: (range: Range) => void,
): void {
  useEffect(() => {
    const element = container.current
    if (element === null) return
    let selection: SelectionConnector | undefined
    let fill: Fill | null = null

    const start = (event: MouseEvent): void => {
      if (current.current.frozenColumns === 0 || !onFillHandle(event)) return
      event.preventDefault()
      event.stopPropagation()
      fill = { source: selected.current, target: selected.current }
    }

    const follow = (event: MouseEvent): void => {
      const position = fill === null ? null : gridPositionOf(event)
      if (fill === null || position === null) return
      event.stopPropagation()
      fill.target = fillTarget(fill.source, position)
      if (selection !== undefined) paint(selection, fill.target, current.current)
    }

    const finish = (event: MouseEvent): void => {
      if (fill === null) return
      event.stopPropagation()
      const { source, target } = fill
      fill = null
      applyChange(fillRange(current.current, source, target))
      setRange(target)
    }

    const unfollow = followSelection(element, (found) => {
      selection = found
    })
    element.addEventListener('mousedown', start, { capture: true })
    window.addEventListener('mousemove', follow, { capture: true })
    window.addEventListener('mouseup', finish, { capture: true })
    return () => {
      unfollow()
      element.removeEventListener('mousedown', start, { capture: true })
      window.removeEventListener('mousemove', follow, { capture: true })
      window.removeEventListener('mouseup', finish, { capture: true })
    }
  }, [container, current, selected, applyChange, setRange])
}
