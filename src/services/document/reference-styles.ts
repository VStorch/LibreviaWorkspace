/**
 * Word only writes `TOC Heading`, `toc 1`…`toc 9` and `caption` when someone uses them: they enter
 * the sheet on insertion, with Word 2013–2021 measures. The internal **name** rules: a Portuguese
 * document calls the `toc 1` style `Sumrio1`, and that is the one used.
 */

import { StyleType, type StyleDefinition, type StyleSheet } from './styles.js'

/** 11 pt, Word's 220 twentieths. */
const TOC_INDENT_MM = 3.88

function findByName(sheet: StyleSheet, name: string): StyleDefinition | undefined {
  const wanted = name.toLowerCase()
  return Object.values(sheet.styles).find(
    (style) => style.type === StyleType.Paragraph && style.name.toLowerCase() === wanted,
  )
}

/** `TOC1`, `TOC1_2`… */
function freeId(sheet: StyleSheet, preferred: string): string {
  if (sheet.styles[preferred] === undefined) return preferred
  let suffix = 2
  while (sheet.styles[`${preferred}_${suffix}`] !== undefined) suffix++
  return `${preferred}_${suffix}`
}

export interface EnsuredStyle {
  readonly sheet: StyleSheet
  readonly id: string
}

function ensure(
  sheet: StyleSheet,
  name: string,
  preferredId: string,
  create: (id: string) => StyleDefinition,
): EnsuredStyle {
  const existing = findByName(sheet, name)
  if (existing !== undefined) return { sheet, id: existing.id }
  const id = freeId(sheet, preferredId)
  return { sheet: { ...sheet, styles: { ...sheet.styles, [id]: create(id) } }, id }
}

const base = {
  type: StyleType.Paragraph,
  hidden: false,
  custom: false,
} as const

export function ensureTocStyle(sheet: StyleSheet, level: number): EnsuredStyle {
  const normal = sheet.defaults.paragraphStyleId ?? undefined
  return ensure(sheet, `toc ${level}`, `TOC${level}`, (id) => ({
    ...base,
    id,
    name: `toc ${level}`,
    qFormat: false,
    uiPriority: 39,
    ...(normal === undefined ? {} : { basedOn: normal, next: normal }),
    paragraph: {
      spaceAfter: 5,
      ...(level > 1 ? { indentMm: Math.round(TOC_INDENT_MM * (level - 1) * 100) / 100 } : {}),
    },
  }))
}

/** `TOC Heading`, which inherits from Heading 1 and stops being a heading (level 9). */
export function ensureTocHeadingStyle(sheet: StyleSheet): EnsuredStyle {
  const heading = findByName(sheet, 'heading 1')?.id
  const normal = sheet.defaults.paragraphStyleId ?? undefined
  return ensure(sheet, 'TOC Heading', 'TOCHeading', (id) => ({
    ...base,
    id,
    name: 'TOC Heading',
    qFormat: true,
    uiPriority: 39,
    ...(heading === undefined ? {} : { basedOn: heading }),
    ...(normal === undefined ? {} : { next: normal }),
    paragraph: { outlineLevel: 9 },
  }))
}

/** `caption`, 9 pt italic, like Word's. */
export function ensureCaptionStyle(sheet: StyleSheet): EnsuredStyle {
  const normal = sheet.defaults.paragraphStyleId ?? undefined
  return ensure(sheet, 'caption', 'Caption', (id) => ({
    ...base,
    id,
    name: 'caption',
    qFormat: true,
    uiPriority: 35,
    ...(normal === undefined ? {} : { basedOn: normal, next: normal }),
    paragraph: { spaceAfter: 10 },
    character: { italic: true, fontSize: '9pt', color: '#44546a' },
  }))
}
