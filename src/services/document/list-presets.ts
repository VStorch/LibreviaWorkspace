/**
 * A biblioteca de listas prontas: as que o Word oferece na galeria de numeração
 * e de marcadores, com as medidas dele.
 *
 * Cada uma são os nove níveis inteiros, e não só o primeiro: Tab dentro da lista
 * desce ao nível seguinte, e é ele que tem de estar definido — senão a sublista
 * sai com a marca padrão, e não com a da lista escolhida.
 */

import { LIST_LEVELS, defaultLevels, type LevelDef } from './list-numbering.js'

export interface ListPreset {
  readonly id: string
  readonly kind: 'bulletList' | 'orderedList'
  /** Chave do catálogo i18n com o nome da lista. */
  readonly labelKey: string
  readonly levels: readonly LevelDef[]
}

const twipsToMm = (twips: number): number => Math.round((twips / 1440) * 25.4 * 100) / 100

const levels = (build: (level: number) => Omit<LevelDef, 'start'> & { start?: number }): LevelDef[] =>
  Array.from({ length: LIST_LEVELS }, (_, level) => ({ start: 1, ...build(level) }))

/** Os formatos que o diálogo oferece, na ordem em que aparecem. */
export const NUMBER_FORMATS = [
  'decimal',
  'decimalZero',
  'lowerLetter',
  'upperLetter',
  'lowerRoman',
  'upperRoman',
  'bullet',
  'none',
] as const

/** As marcas que o diálogo oferece — todas com glifo conhecido do Word (`ListLevels.cs`). */
export const BULLET_MARKS = ['•', 'o', '▪', '➢', '●', '□', '✓', '–'] as const

export const LIST_PRESETS: readonly ListPreset[] = [
  {
    id: 'outline',
    kind: 'orderedList',
    labelKey: 'document.lists.presetOutline',
    levels: defaultLevels('orderedList'),
  },
  {
    // "1 1.1 1.1.1" do Word: cada nível leva os de cima, e o recuo pendente
    // cresce com o número, que fica mais largo a cada nível.
    id: 'legal',
    kind: 'orderedList',
    labelKey: 'document.lists.presetLegal',
    levels: levels((level) => ({
      fmt: 'decimal',
      text: Array.from({ length: level + 1 }, (_, index) => `%${index + 1}.`).join(''),
      indentMm: twipsToMm([360, 792, 1224, 1728, 2232, 2736, 3240, 3744, 4320][level]!),
      hangingMm: twipsToMm([360, 432, 504, 648, 792, 936, 1080, 1224, 1440][level]!),
    })),
  },
  {
    id: 'roman',
    kind: 'orderedList',
    labelKey: 'document.lists.presetRoman',
    levels: levels((level) => ({
      fmt: ['upperRoman', 'upperLetter', 'decimal', 'lowerLetter', 'lowerRoman'][level % 5]!,
      text: `%${level + 1}.`,
      indentMm: twipsToMm(720 * (level + 1)),
      hangingMm: twipsToMm(360),
    })),
  },
  {
    id: 'paren',
    kind: 'orderedList',
    labelKey: 'document.lists.presetParen',
    levels: levels((level) => ({
      fmt: ['decimal', 'lowerLetter', 'lowerRoman'][level % 3]!,
      text: `%${level + 1})`,
      indentMm: twipsToMm(720 * (level + 1)),
      hangingMm: twipsToMm(360),
    })),
  },
  {
    id: 'bullets',
    kind: 'bulletList',
    labelKey: 'document.lists.presetBullets',
    levels: defaultLevels('bulletList'),
  },
  {
    id: 'arrows',
    kind: 'bulletList',
    labelKey: 'document.lists.presetArrows',
    levels: levels((level) => ({
      fmt: 'bullet',
      text: ['➢', '▪', '•'][level % 3]!,
      indentMm: twipsToMm(720 * (level + 1)),
      hangingMm: twipsToMm(360),
    })),
  },
  {
    id: 'dashes',
    kind: 'bulletList',
    labelKey: 'document.lists.presetDashes',
    levels: levels((level) => ({
      fmt: 'bullet',
      text: '–',
      indentMm: twipsToMm(720 * (level + 1)),
      hangingMm: twipsToMm(360),
    })),
  },
]

/** O tipo de nó que uma definição pede, pelo primeiro nível — como o leitor decide. */
export function kindOfLevels(list: readonly LevelDef[]): 'bulletList' | 'orderedList' {
  const fmt = list[0]?.fmt
  return fmt === 'bullet' || fmt === 'none' ? 'bulletList' : 'orderedList'
}
