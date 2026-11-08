import { describe, expect, it } from 'vitest'
import {
  canOpenAnother,
  canPlaceTab,
  closeTab,
  comparablePath,
  findOpenTab,
  HOME_ONLY,
  HOME_TAB_ID,
  MAX_OPEN_TABS,
  moveTab,
  nextTab,
  openTab,
  placeTab,
  previousTab,
  type TabOrder,
} from './tab-order.js'

const order = (ids: string[], activeId: string): TabOrder => ({ ids, activeId })

describe('openTab', () => {
  it('appends the new tab after the others and activates it, keeping Home', () => {
    expect(openTab(HOME_ONLY, 'a')).toEqual(order([HOME_TAB_ID, 'a'], 'a'))
  })

  it('activates a tab that is already open instead of adding it twice', () => {
    expect(openTab(order(['a', 'b'], 'b'), 'a')).toEqual(order(['a', 'b'], 'a'))
  })
})

describe('placeTab', () => {
  it('puts the new tab where the replaced one was and activates it', () => {
    expect(placeTab(order(['a', 'b', 'c'], 'b'), 'n', 'b')).toEqual(order(['a', 'n', 'c'], 'n'))
  })

  it('replaces Home when Home is the one given', () => {
    expect(placeTab(HOME_ONLY, 'n', HOME_TAB_ID)).toEqual(order(['n'], 'n'))
  })

  it('appends at the end when nothing is replaced', () => {
    expect(placeTab(order(['a'], 'a'), 'n', null)).toEqual(order(['a', 'n'], 'n'))
  })
})

describe('canPlaceTab', () => {
  const full = order(
    Array.from({ length: MAX_OPEN_TABS }, (_, index) => `t${index}`),
    't0',
  )

  it('lets a file tab be swapped even when the limit is reached', () => {
    expect(canPlaceTab(full, 't0')).toBe(true)
  })

  it('refuses a new tab, or one replacing Home, when the limit is reached', () => {
    expect(canPlaceTab(full, null)).toBe(false)
    expect(canPlaceTab({ ids: [...full.ids, HOME_TAB_ID], activeId: HOME_TAB_ID }, HOME_TAB_ID)).toBe(false)
  })
})

describe('closeTab', () => {
  it('activates the right neighbour when the active tab closes', () => {
    expect(closeTab(order(['a', 'b', 'c'], 'b'), 'b')).toEqual(order(['a', 'c'], 'c'))
  })

  it('activates the left neighbour when the active tab was the last one', () => {
    expect(closeTab(order(['a', 'b', 'c'], 'c'), 'c')).toEqual(order(['a', 'b'], 'b'))
  })

  it('keeps the active tab when an inactive tab closes', () => {
    expect(closeTab(order(['a', 'b', 'c'], 'a'), 'c')).toEqual(order(['a', 'b'], 'a'))
  })

  it('leaves Home alone and active when the only document closes', () => {
    expect(closeTab(order(['a'], 'a'), 'a')).toEqual(HOME_ONLY)
  })

  it('falls back to Home when the last document next to Home closes', () => {
    expect(closeTab(order([HOME_TAB_ID, 'a'], 'a'), 'a')).toEqual(HOME_ONLY)
  })

  it('lets Home close while documents stay open', () => {
    expect(closeTab(order([HOME_TAB_ID, 'a'], HOME_TAB_ID), HOME_TAB_ID)).toEqual(order(['a'], 'a'))
  })

  it('keeps Home when it is the only tab', () => {
    expect(closeTab(HOME_ONLY, HOME_TAB_ID)).toEqual(HOME_ONLY)
  })

  it('ignores an unknown id', () => {
    const tabs = order(['a', 'b'], 'a')
    expect(closeTab(tabs, 'z')).toBe(tabs)
  })
})

describe('canOpenAnother', () => {
  const withDocuments = (count: number, home: boolean): TabOrder => {
    const ids = Array.from({ length: count }, (_, i) => `t${i}`)
    return order(home ? [HOME_TAB_ID, ...ids] : ids, ids[0] ?? HOME_TAB_ID)
  }

  it('allows opening up to the limit', () => {
    expect(canOpenAnother(withDocuments(MAX_OPEN_TABS - 1, false))).toBe(true)
  })

  it('refuses at the limit', () => {
    expect(canOpenAnother(withDocuments(MAX_OPEN_TABS, false))).toBe(false)
  })

  it('does not count Home against the limit', () => {
    expect(canOpenAnother(withDocuments(MAX_OPEN_TABS - 1, true))).toBe(true)
  })
})

describe('nextTab and previousTab', () => {
  const tabs = order(['a', 'b', 'c'], 'b')

  it('move one tab right and left', () => {
    expect(nextTab(tabs).activeId).toBe('c')
    expect(previousTab(tabs).activeId).toBe('a')
  })

  it('wrap around at both ends', () => {
    expect(nextTab(order(['a', 'b', 'c'], 'c')).activeId).toBe('a')
    expect(previousTab(order(['a', 'b', 'c'], 'a')).activeId).toBe('c')
  })

  it('stay put with a single tab', () => {
    expect(nextTab(HOME_ONLY)).toEqual(HOME_ONLY)
    expect(previousTab(HOME_ONLY)).toEqual(HOME_ONLY)
  })
})

describe('moveTab', () => {
  const tabs = order(['a', 'b', 'c', 'd'], 'b')

  it('moves a tab to the given index and keeps the active tab', () => {
    expect(moveTab(tabs, 'a', 2)).toEqual(order(['b', 'c', 'a', 'd'], 'b'))
    expect(moveTab(tabs, 'd', 0)).toEqual(order(['d', 'a', 'b', 'c'], 'b'))
  })

  it('clamps an index outside the bar', () => {
    expect(moveTab(tabs, 'b', 99).ids).toEqual(['a', 'c', 'd', 'b'])
    expect(moveTab(tabs, 'c', -3).ids).toEqual(['c', 'a', 'b', 'd'])
  })

  it('ignores an unknown id', () => {
    expect(moveTab(tabs, 'z', 0)).toBe(tabs)
  })
})

describe('findOpenTab', () => {
  const tabs = [
    { id: HOME_TAB_ID, path: null },
    { id: 'new', path: null },
    { id: 'report', path: '/home/ana/Report.docx' },
    { id: 'budget', path: 'C:\\Users\\Ana\\Budget.xlsx' },
  ]

  it('finds the tab with the same path', () => {
    expect(findOpenTab(tabs, '/home/ana/Report.docx', 'linux')).toBe('report')
  })

  it('ignores repeated separators', () => {
    expect(findOpenTab(tabs, '/home//ana/Report.docx', 'linux')).toBe('report')
  })

  it('is case-sensitive outside Windows', () => {
    expect(findOpenTab(tabs, '/home/ana/report.docx', 'linux')).toBeNull()
    expect(findOpenTab(tabs, '/home/ana/report.docx', 'darwin')).toBeNull()
  })

  it('on Windows ignores case and treats both separators alike', () => {
    expect(findOpenTab(tabs, 'c:/users/ana/BUDGET.xlsx', 'win32')).toBe('budget')
  })

  it('never matches unsaved tabs', () => {
    expect(findOpenTab(tabs, '', 'linux')).toBeNull()
  })
})

describe('comparablePath', () => {
  it('drops trailing separators but keeps the root', () => {
    expect(comparablePath('/home/ana/', 'linux')).toBe('/home/ana')
    expect(comparablePath('/', 'linux')).toBe('/')
    expect(comparablePath('C:\\', 'win32')).toBe('c:/')
  })

  it('keeps a backslash as a file name character outside Windows', () => {
    expect(comparablePath('/tmp/a\\b', 'linux')).toBe('/tmp/a\\b')
  })

  it('keeps the double slash that starts a Windows network path', () => {
    expect(comparablePath('\\\\Server\\Share\\File.docx', 'win32')).toBe('//server/share/file.docx')
  })
})
