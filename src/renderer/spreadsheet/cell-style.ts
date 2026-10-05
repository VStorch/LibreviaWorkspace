import { getCell, type Sheet } from '@services/spreadsheet/model.js'

/**
 * At draw time: scrolling recreates the cells, and a style kept in the DOM would vanish with them.
 */
export function cellStyleOf(sheet: Sheet, row: number, column: number): Record<string, string> {
  const style = getCell(sheet, row, column)?.style
  if (style === undefined) return {}

  const css: Record<string, string> = {}
  if (style.bold === true) css['fontWeight'] = '700'
  if (style.italic === true) css['fontStyle'] = 'italic'
  if (style.underline === true) css['textDecoration'] = 'underline'
  if (style.color !== undefined) css['color'] = style.color
  if (style.background !== undefined) css['backgroundColor'] = style.background
  if (style.align !== undefined) css['textAlign'] = style.align

  for (const side of style.borders ?? []) {
    css[`border${side[0]!.toUpperCase()}${side.slice(1)}`] = '1px solid #555'
  }

  return css
}
