/**
 * The name is what the table of contents and references cite; the `w:id` pairs both ends and is
 * unique.
 */

export const BOOKMARK_NAME_MAX = 40

/**
 * As in Word: a letter, then letters, digits and underscores, up to 40. No leading underscore,
 * which marks the hidden ones (`_Toc…`, `_Ref…`).
 */
export function isValidBookmarkName(name: string): boolean {
  return new RegExp(`^\\p{L}[\\p{L}\\p{N}_]{0,${BOOKMARK_NAME_MAX - 1}}$`, 'u').test(name)
}

/** The ones Word creates for the table of contents and for references. */
export function isHiddenBookmark(name: string): boolean {
  return name.startsWith('_')
}

/** `w:id` is an integer in the schema; anything else is not counted. */
export function nextBookmarkId(existing: Iterable<string>): string {
  let highest = -1
  for (const id of existing) {
    const value = Number(id)
    if (Number.isInteger(value) && value > highest) highest = value
  }
  return String(highest + 1)
}

/** Word draws random digits; here they count up from the highest, without relying on luck. */
export function hiddenBookmarkName(prefix: '_Ref' | '_Toc', existing: Iterable<string>): string {
  let highest = 0
  const pattern = new RegExp(`^${prefix}(\\d+)$`)
  for (const name of existing) {
    const match = pattern.exec(name)
    if (match !== null) highest = Math.max(highest, Number(match[1]))
  }
  return `${prefix}${String(highest + 1).padStart(9, '0')}`
}
