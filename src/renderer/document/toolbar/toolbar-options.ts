/** A lista de fontes não mora aqui: depende da máquina e do documento (`useFontFamilies`). */

import type { MessageKey } from '@shared/i18n/index.js'

/** Os do Word. */
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

/** Em linhas do Word, e não em CSS; vazio é "Simples". */
export function lineHeights(t: (key: MessageKey) => string): readonly { value: string; label: string }[] {
  return [
    { value: '', label: t('document.paragraph.spacingSingle') },
    { value: '1.15', label: '1,15' },
    { value: '1.5', label: '1,5' },
    { value: '2', label: t('document.paragraph.spacingDouble') },
  ]
}

/** Um `<select>` sem a opção do próprio valor mostra a primeira, e a barra mentiria. */
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
