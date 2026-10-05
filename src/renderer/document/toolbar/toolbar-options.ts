/**
 * The font list does not live here: it depends on the machine and the document (`useFontFamilies`).
 */

import type { MessageKey } from '@shared/i18n/index.js'

/** Word's. */
export const FONT_SIZES = [
  '8',
  '9',
  '10',
  '11',
  '12',
  '14',
  '16',
  '18',
  '20',
  '24',
  '28',
  '32',
  '36',
  '48',
  '72',
] as const

/** In Word lines, not CSS; empty means "Single". */
export function lineHeights(t: (key: MessageKey) => string): readonly { value: string; label: string }[] {
  return [
    { value: '', label: t('document.paragraph.spacingSingle') },
    { value: '1.15', label: '1,15' },
    { value: '1.5', label: '1,5' },
    { value: '2', label: t('document.paragraph.spacingDouble') },
  ]
}

/**
 * A `<select>` without its own value as an option shows the first one, and the toolbar would lie.
 */
export function withCurrent(
  options: readonly { readonly value: string; readonly label: string }[],
  current: string,
  label: (value: string) => string = (value) => value,
): readonly { value: string; label: string }[] {
  if (current === '' || options.some((option) => option.value === current)) {
    return options as readonly { value: string; label: string }[]
  }

  return [...options, { value: current, label: label(current) }]
}
