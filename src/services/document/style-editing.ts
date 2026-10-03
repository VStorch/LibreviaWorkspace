/**
 * Toda função devolve uma folha nova: é pela referência que o CSS se regera.
 * `StyleWriter.cs` só mexe no `w:style` que mudou. O formulário mostra o que o
 * estilo vale e grava só o campo mudado, para mudar o pai continuar alcançando o
 * filho.
 */

import { firstFontOf } from './line-metrics.js'
import {
  FirstLineKind,
  LineSpacingKind,
  paragraphDraftFrom,
  type ParagraphDraft,
} from './paragraph-format.js'
import { resolveCharacterStyle, resolveStyle, styleAttrsOf } from './style-cascade.js'
import {
  StyleType,
  type LineSpacing,
  type StyleCharacterFormat,
  type StyleDefinition,
  type StyleParagraphFormat,
  type StyleSheet,
} from './styles.js'

/** O nível de título de um estilo, pelo nome interno — `heading 1` em qualquer idioma. */
export function headingLevelOfStyle(style: StyleDefinition | undefined): number | null {
  const match = /^heading ([1-6])$/i.exec(style?.name ?? '')
  return match === null ? null : Number(match[1])
}

export function withStyle(sheet: StyleSheet, style: StyleDefinition): StyleSheet {
  return { ...sheet, styles: { ...sheet.styles, [style.id]: style } }
}

/** Como o Word: sem acento, espaço nem pontuação, e um número no fim quando já existe. */
export function newStyleId(sheet: StyleSheet, name: string): string {
  const base =
    name
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .replace(/[^A-Za-z0-9]/g, '') || 'Estilo'
  let id = base
  for (let suffix = 2; id in sheet.styles; suffix++) id = `${base}${suffix}`
  return id
}

export interface NewStyle {
  readonly name: string
  readonly type: StyleType
  readonly basedOn?: string | undefined
  readonly next?: string | undefined
}

/** Um estilo personalizado, vazio: tudo o que ele vale vem do pai. */
export function createStyle(sheet: StyleSheet, input: NewStyle): { sheet: StyleSheet; id: string } {
  const id = newStyleId(sheet, input.name)
  const style: StyleDefinition = {
    id,
    name: input.name.trim(),
    type: input.type,
    qFormat: true,
    hidden: false,
    custom: true,
    ...(input.basedOn === undefined ? {} : { basedOn: input.basedOn }),
    ...(input.next === undefined || input.type !== StyleType.Paragraph ? {} : { next: input.next }),
  }
  return { sheet: withStyle(sheet, style), id }
}

/** O nome do estilo embutido (`heading 1`) não muda: é como o Word reconhece o título. */
export function withIdentity(
  sheet: StyleSheet,
  id: string,
  identity: {
    readonly name: string
    readonly basedOn?: string | undefined
    readonly next?: string | undefined
  },
): StyleSheet {
  const style = sheet.styles[id]
  if (style === undefined) return sheet

  const rest: StyleDefinition = { ...style }
  delete (rest as { basedOn?: string }).basedOn
  delete (rest as { next?: string }).next
  const name = style.custom && identity.name.trim() !== '' ? identity.name.trim() : style.name
  // Herdar de si mesmo seria laço, e o resolvedor cortaria a cadeia.
  const basedOn = identity.basedOn === id ? style.basedOn : identity.basedOn
  return withStyle(sheet, {
    ...rest,
    name,
    ...(basedOn === undefined ? {} : { basedOn }),
    ...(identity.next === undefined ? {} : { next: identity.next }),
  })
}

export interface StyleDraft {
  readonly paragraph: ParagraphDraft
  /** Nome da fonte, ou vazio quando ninguém na cadeia diz. */
  readonly fontFamily: string
  /** Pontos, ou `null` quando ninguém na cadeia diz. */
  readonly fontSize: number | null
  readonly bold: boolean
  readonly italic: boolean
  readonly underline: boolean
}

export function styleDraftOf(sheet: StyleSheet, id: string): StyleDraft {
  const style = sheet.styles[id]
  const character =
    style?.type === StyleType.Character ? resolveCharacterStyle(sheet, id) : resolveStyle(sheet, id).character
  const paragraph = paragraphDraftFrom(styleAttrsOf(resolveStyle(sheet, id)))
  return {
    paragraph,
    fontFamily: firstFontOf(character.fontFamily) ?? '',
    fontSize: pointsOf(character.fontSize),
    bold: character.bold === true,
    italic: character.italic === true,
    underline: character.underline === true,
  }
}

/** Campo igual continua como estava, declarado ou herdado. */
export function styleWithDraft(sheet: StyleSheet, id: string, draft: StyleDraft): StyleSheet {
  const style = sheet.styles[id]
  if (style === undefined) return sheet

  const shown = styleDraftOf(sheet, id)
  const before = shown.paragraph
  const after = draft.paragraph
  const paragraph: Record<string, unknown> = { ...style.paragraph }
  const character: Record<string, unknown> = { ...style.character }

  if (style.type === StyleType.Paragraph) {
    if (after.align !== before.align) paragraph['textAlign'] = after.align
    if (after.spaceBefore !== before.spaceBefore) paragraph['spaceBefore'] = after.spaceBefore
    if (after.spaceAfter !== before.spaceAfter) paragraph['spaceAfter'] = after.spaceAfter
    if (
      after.lineSpacingKind !== before.lineSpacingKind ||
      after.lineSpacingValue !== before.lineSpacingValue
    ) {
      paragraph['lineSpacing'] = lineSpacingOf(after)
    }
    if (after.indentLeftMm !== before.indentLeftMm) paragraph['indentMm'] = after.indentLeftMm
    if (after.indentRightMm !== before.indentRightMm) paragraph['indentRightMm'] = after.indentRightMm
    if (after.firstLineKind !== before.firstLineKind || after.firstLineMm !== before.firstLineMm) {
      paragraph['firstLineMm'] =
        after.firstLineKind === FirstLineKind.None
          ? 0
          : after.firstLineKind === FirstLineKind.Hanging
            ? -Math.abs(after.firstLineMm)
            : Math.abs(after.firstLineMm)
    }
    if (after.keepNext !== before.keepNext) paragraph['keepNext'] = after.keepNext
  }

  if (draft.fontFamily !== shown.fontFamily && draft.fontFamily !== '')
    character['fontFamily'] = draft.fontFamily
  if (draft.fontSize !== shown.fontSize && draft.fontSize !== null)
    character['fontSize'] = `${draft.fontSize}pt`
  if (draft.bold !== shown.bold) character['bold'] = draft.bold
  if (draft.italic !== shown.italic) character['italic'] = draft.italic
  if (draft.underline !== shown.underline) character['underline'] = draft.underline

  const rest: StyleDefinition = { ...style }
  delete (rest as { paragraph?: unknown }).paragraph
  delete (rest as { character?: unknown }).character
  return withStyle(sheet, {
    ...rest,
    ...(Object.keys(paragraph).length === 0 ? {} : { paragraph: paragraph as StyleParagraphFormat }),
    ...(Object.keys(character).length === 0 ? {} : { character: character as StyleCharacterFormat }),
  })
}

/**
 * "Atualizar estilo a partir da seleção", só quando a pessoa pede.
 *
 * @param effective os atributos do bloco com o estilo por baixo (`effectiveAttrs`).
 */
export function styleFromBlock(
  sheet: StyleSheet,
  id: string,
  effective: Record<string, unknown>,
): StyleDraft {
  const shown = styleDraftOf(sheet, id)
  const family = typeof effective['fontFamily'] === 'string' ? firstFontOf(effective['fontFamily']) : null
  const size = typeof effective['fontSize'] === 'string' ? pointsOf(effective['fontSize']) : null
  return {
    ...shown,
    paragraph: paragraphDraftFrom(effective),
    fontFamily: family ?? shown.fontFamily,
    fontSize: size ?? shown.fontSize,
  }
}

/** O estilo que vem depois de um parágrafo deste, ao dar Enter no fim dele. */
export function nextStyleIdOf(sheet: StyleSheet, id: string | null): string | null {
  const own = id ?? sheet.defaults.paragraphStyleId
  if (own === null) return null
  const style = sheet.styles[own]
  if (style === undefined) return own
  return style.next !== undefined && style.next in sheet.styles ? style.next : own
}

function lineSpacingOf(draft: ParagraphDraft): LineSpacing {
  if (draft.lineSpacingKind === LineSpacingKind.AtLeast)
    return { kind: 'atLeast', pt: draft.lineSpacingValue }
  if (draft.lineSpacingKind === LineSpacingKind.Multiple) {
    return { kind: 'multiple', factor: draft.lineSpacingValue }
  }
  return { kind: 'multiple', factor: 1 }
}

function pointsOf(size: string | undefined): number | null {
  if (size === undefined || !size.endsWith('pt')) return null
  const points = Number(size.slice(0, -2))
  return Number.isFinite(points) && points > 0 ? points : null
}
