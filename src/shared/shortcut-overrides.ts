import { SHORTCUTS, canonicalKeyOf, type Shortcut, type ShortcutId, type ShortcutKey } from './shortcuts.js'

export type ShortcutTable = Readonly<Record<ShortcutId, Shortcut>>

/** What the person chose, by shortcut id, in the form `keyText` writes; the rest keep the default. */
export type ShortcutOverrides = Readonly<Record<string, string>>

/**
 * Keys the table leaves to Chromium and Electron roles (copy, paste, undo, quit, full screen) or to
 * Tiptap's alignment: moving a command onto one would take it from the editor without a word.
 */
const RESERVED = new Set(
  [
    ...['C', 'X', 'V', 'A', 'Z', 'Y', 'Q'].map((key) => ({ mod: true, key }) as const),
    ...['Z', 'L', 'E', 'R', 'J', 'I'].map((key) => ({ mod: true, shift: true, key }) as const),
    { key: 'F11' },
  ].map(canonicalKeyOf),
)

const FUNCTION_KEY = /^F([1-9]|1\d|2[0-4])$/u

export const ShortcutProblem = {
  NeedsModifier: 'needs-modifier',
  Reserved: 'reserved',
  Taken: 'taken',
} as const

export type ShortcutProblem =
  | { readonly kind: typeof ShortcutProblem.NeedsModifier }
  | { readonly kind: typeof ShortcutProblem.Reserved }
  | { readonly kind: typeof ShortcutProblem.Taken; readonly by: ShortcutId }

/** Keys the page hears itself, or the developer's reload. */
const PAGE_HANDLED = new Set<string>(['reload', 'nextSheet', 'previousSheet'])

/** Menu keys only: the editor's are registered by its keymap or by Tiptap, the sheet's by the page. */
export function isRemappable(id: string): id is ShortcutId {
  if (!Object.hasOwn(SHORTCUTS, id)) return false
  const shortcut: Shortcut = SHORTCUTS[id as ShortcutId]
  return shortcut.owner === 'menu' && shortcut.registeredBy === undefined && !PAGE_HANDLED.has(id)
}

/** `Mod+Shift+K`: the stored form, which keeps the key's case so `PageDown` stays `PageDown`. */
export function keyText(key: ShortcutKey): string {
  return [
    ...(key.mod === true ? ['Mod'] : []),
    ...(key.shift === true ? ['Shift'] : []),
    ...(key.alt === true ? ['Alt'] : []),
    key.key,
  ].join('+')
}

export function parseKeyText(text: string): ShortcutKey | null {
  const parts = text.split('+')
  const name = parts.pop()
  if (name === undefined || name === '') return null
  if (parts.some((modifier) => !['Mod', 'Shift', 'Alt'].includes(modifier))) return null
  return {
    ...(parts.includes('Mod') ? { mod: true } : {}),
    ...(parts.includes('Shift') ? { shift: true } : {}),
    ...(parts.includes('Alt') ? { alt: true } : {}),
    key: name,
  }
}

export function shortcutProblem(
  table: ShortcutTable,
  id: ShortcutId,
  key: ShortcutKey,
): ShortcutProblem | null {
  if (key.mod !== true && !FUNCTION_KEY.test(key.key)) return { kind: ShortcutProblem.NeedsModifier }
  const canonical = canonicalKeyOf(key)
  if (RESERVED.has(canonical)) return { kind: ShortcutProblem.Reserved }
  const holder = (Object.keys(table) as ShortcutId[]).find(
    (other) => other !== id && canonicalKeyOf(table[other].key) === canonical,
  )
  return holder === undefined ? null : { kind: ShortcutProblem.Taken, by: holder }
}

function withKeys(base: ShortcutTable, keys: ReadonlyMap<ShortcutId, ShortcutKey>): ShortcutTable {
  const table: Record<string, Shortcut> = { ...base }
  for (const [id, key] of keys) table[id] = { ...base[id], key }
  return table as ShortcutTable
}

/**
 * All at once, so two commands can swap keys; then any choice that clashes, which only a hand-edited
 * file can hold, goes back to its default until nothing clashes.
 */
export function applyShortcutOverrides(base: ShortcutTable, overrides: ShortcutOverrides): ShortcutTable {
  const keys = new Map<ShortcutId, ShortcutKey>()
  for (const [id, text] of Object.entries(overrides)) {
    const key = parseKeyText(text)
    if (isRemappable(id) && key !== null) keys.set(id, key)
  }
  let table = withKeys(base, keys)
  for (;;) {
    const clashing = [...keys.entries()].filter(([id, key]) => shortcutProblem(table, id, key) !== null)
    if (clashing.length === 0) return table
    for (const [id] of clashing) keys.delete(id)
    table = withKeys(base, keys)
  }
}

const NAMED_CODES: Readonly<Record<string, string>> = {
  Minus: '-',
  Equal: '=',
  BracketLeft: '[',
  BracketRight: ']',
  Semicolon: ';',
  Quote: "'",
  Comma: ',',
  Period: '.',
  Slash: '/',
  Backslash: '\\',
  Backquote: '`',
  ArrowUp: 'Up',
  ArrowDown: 'Down',
  ArrowLeft: 'Left',
  ArrowRight: 'Right',
  NumpadAdd: 'numadd',
  NumpadSubtract: 'numsub',
  NumpadMultiply: 'nummult',
  NumpadDivide: 'numdiv',
  NumpadDecimal: 'numdec',
}

const PLAIN_CODES = new Set([
  'Enter',
  'Tab',
  'Space',
  'Backspace',
  'Delete',
  'Insert',
  'Home',
  'End',
  'PageUp',
  'PageDown',
])

/** By the physical key, as Electron names it: with Shift, `key` would say `!` where the key is `1`. */
function keyNameOf(code: string): string | null {
  const letterOrDigit = /^(?:Key([A-Z])|Digit(\d)|Numpad(\d))$/u.exec(code)
  if (letterOrDigit !== null) {
    const [, letter, digit, pad] = letterOrDigit
    return letter ?? digit ?? `num${pad}`
  }
  if (FUNCTION_KEY.test(code) || PLAIN_CODES.has(code)) return code
  return NAMED_CODES[code] ?? null
}

export interface KeyPress {
  readonly code: string
  readonly ctrlKey: boolean
  readonly metaKey: boolean
  readonly shiftKey: boolean
  readonly altKey: boolean
}

/** Null for a modifier alone, or a key Electron has no name for. */
export function shortcutKeyOf(press: KeyPress): ShortcutKey | null {
  const name = keyNameOf(press.code)
  if (name === null) return null
  return {
    ...(press.ctrlKey || press.metaKey ? { mod: true } : {}),
    ...(press.shiftKey ? { shift: true } : {}),
    ...(press.altKey ? { alt: true } : {}),
    key: name,
  }
}
