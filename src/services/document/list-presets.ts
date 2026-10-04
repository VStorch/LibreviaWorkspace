/**
 * As listas da galeria do Word, com as medidas dele. Os nove níveis, porque Tab
 * desce ao nível seguinte, que tem de estar definido.
 */

import { twipsToMm } from '@services/units.js'
import { LIST_LEVELS, defaultLevels, type LevelDef } from './list-numbering.js'

export interface ListPreset {
  readonly id: string
  readonly kind: 'bulletList' | 'orderedList'
  readonly labelKey: string
  readonly levels: readonly LevelDef[]
}

const twipsToRoundedMm = (twips: number): number => Math.round(twipsToMm(twips) * 100) / 100

const levels = (build: (level: number) => Omit<LevelDef, 'start'> & { start?: number }): LevelDef[] =>
  Array.from({ length: LIST_LEVELS }, (_, level) => ({ start: 1, ...build(level) }))

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
      indentMm: twipsToRoundedMm([360, 792, 1224, 1728, 2232, 2736, 3240, 3744, 4320][level]!),
      hangingMm: twipsToRoundedMm([360, 432, 504, 648, 792, 936, 1080, 1224, 1440][level]!),
    })),
  },
  {
    id: 'roman',
    kind: 'orderedList',
    labelKey: 'document.lists.presetRoman',
    levels: levels((level) => ({
      fmt: ['upperRoman', 'upperLetter', 'decimal', 'lowerLetter', 'lowerRoman'][level % 5]!,
      text: `%${level + 1}.`,
      indentMm: twipsToRoundedMm(720 * (level + 1)),
      hangingMm: twipsToRoundedMm(360),
    })),
  },
  {
    id: 'paren',
    kind: 'orderedList',
    labelKey: 'document.lists.presetParen',
    levels: levels((level) => ({
      fmt: ['decimal', 'lowerLetter', 'lowerRoman'][level % 3]!,
      text: `%${level + 1})`,
      indentMm: twipsToRoundedMm(720 * (level + 1)),
      hangingMm: twipsToRoundedMm(360),
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
      indentMm: twipsToRoundedMm(720 * (level + 1)),
      hangingMm: twipsToRoundedMm(360),
    })),
  },
  {
    id: 'dashes',
    kind: 'bulletList',
    labelKey: 'document.lists.presetDashes',
    levels: levels((level) => ({
      fmt: 'bullet',
      text: '–',
      indentMm: twipsToRoundedMm(720 * (level + 1)),
      hangingMm: twipsToRoundedMm(360),
    })),
  },
]

/** Pelo primeiro nível, como o leitor decide. */
export function kindOfLevels(list: readonly LevelDef[]): 'bulletList' | 'orderedList' {
  const fmt = list[0]?.fmt
  return fmt === 'bullet' || fmt === 'none' ? 'bulletList' : 'orderedList'
}
