/**
 * O Word só grava `TOC Heading`, `toc 1`…`toc 9` e `caption` quando alguém os
 * usa: eles entram na folha na hora de inserir, com as medidas do Word
 * 2013–2021. Manda o **nome** interno: o documento em português chama de
 * `Sumrio1` o estilo `toc 1`, e é ele que se usa.
 */

import { StyleType, type StyleDefinition, type StyleSheet } from './styles.js'

/** 11 pt, os 220 vinte-avos do Word. */
const TOC_INDENT_MM = 3.88

function findByName(sheet: StyleSheet, name: string): StyleDefinition | undefined {
  const wanted = name.toLowerCase()
  return Object.values(sheet.styles).find(
    (style) => style.type === StyleType.Paragraph && style.name.toLowerCase() === wanted,
  )
}

/** Um id livre a partir do preferido: `TOC1`, `TOC1_2`… */
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

/** O título do sumário: `TOC Heading`, que herda do Título 1 e deixa de ser título (nível 9). */
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

/** A legenda: `caption`, em itálico de 9 pt, como a do Word. */
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
