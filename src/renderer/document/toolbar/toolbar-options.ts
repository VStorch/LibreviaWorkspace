/**
 * As listas fechadas da barra do documento.
 *
 * Fora do componente porque são dados, e porque a barra já tem trabalho
 * suficiente: o arquivo dela passava de trezentas linhas e continua crescendo a
 * cada recurso. A lista de fontes **não** mora aqui — ela depende da máquina e
 * do documento aberto, e vem de `useFontFamilies`.
 */

import type { MessageKey } from '@shared/i18n/index.js'

/** Os corpos que o Word oferece na caixa de tamanho. */
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

/**
 * A entrelinha do seletor rápido, em linhas do Word.
 *
 * Vazio é "Simples" — a altura que a própria fonte pede. Os números são fatores
 * de linha, como o Word os mostra, e não a medida do CSS: 1,5 linha em Calibri é
 * `line-height: 1.8311`, e quem faz a conta é `paragraph-format`. O resto das
 * opções mora no diálogo de parágrafo.
 */
export function lineHeights(t: (key: MessageKey) => string): readonly { value: string; label: string }[] {
  return [
    { value: '', label: t('document.paragraph.spacingSingle') },
    { value: '1.15', label: '1,15' },
    { value: '1.5', label: '1,5' },
    { value: '2', label: t('document.paragraph.spacingDouble') },
  ]
}

export function blockStyles(t: (key: MessageKey) => string): readonly { value: string; label: string }[] {
  return [
    { value: 'paragraph', label: t('document.styleAndFont.normalText') },
    { value: '1', label: t('document.styleAndFont.heading1') },
    { value: '2', label: t('document.styleAndFont.heading2') },
    { value: '3', label: t('document.styleAndFont.heading3') },
    { value: '4', label: t('document.styleAndFont.heading4') },
  ]
}

/**
 * As opções de um seletor, com o valor atual dentro dela quando faltar.
 *
 * Um `<select>` cujo valor não está entre as opções mostra a **primeira** da
 * lista: a barra passaria a dizer "Simples" num parágrafo de entrelinha 14 pt, e
 * "8" num texto de 10,5 pt. O documento importado traz medidas que nenhuma lista
 * fechada prevê, e a barra tem de contá-las, não escondê-las.
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
