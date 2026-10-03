/**
 * Uma regra por estilo de parágrafo, com **todo** campo explícito (zero quando
 * a cadeia cala), para não herdar a margem do `p` ou o negrito do `h1`. A
 * formatação direta sai inline e vence qualquer seletor. Só os filhos diretos de
 * `.page__content`: lista e célula têm regras próprias em `content-styles.ts`.
 *
 * `w:contextualSpacing` fica de fora: a paginação conferida contra o corpus foi
 * medida sem ele, e ele volta intacto ao arquivo.
 */

import { cssLineHeightOf, explicitCssLineHeightOf } from './line-metrics.js'
import {
  headingStyleOf,
  resolveCharacterStyle,
  resolveStyle,
  usableFactor,
  type ResolvedStyle,
} from './style-cascade.js'
import { StyleType, type StyleCharacterFormat, type StyleSheet } from './styles.js'

const HEADING_LEVELS = 6

export function styleSheetCss(sheet: StyleSheet): string {
  const base = resolveStyle(sheet, null)
  const rules = [
    `.page__content { ${declarations(characterCss(base)).join(' ')} }`,
    // O espaço entre blocos que não são parágrafo é o antes do estilo padrão.
    `.page__content > * + * { margin-top: ${points(base.paragraph.spaceBefore)}; }`,
    rule('.page__content > p:not([data-style-id])', base),
  ]

  for (let level = 1; level <= HEADING_LEVELS; level++) {
    const style = headingStyleOf(sheet, level)
    rules.push(rule(`.page__content > h${level}:not([data-style-id])`, style))
  }

  // O id que o documento não define: o Word o desenha só com os padrões. Antes
  // das regras por id, que vencem por virem depois.
  rules.push(rule('.page__content > [data-style-id]', resolveStyle(sheet, '')))

  // Os de caractere, em qualquer profundidade, só com o que a cadeia declara.
  for (const style of Object.values(sheet.styles)) {
    if (style.type !== StyleType.Character) continue
    const css = declaredCharacterCss(resolveCharacterStyle(sheet, style.id))
    if (css.length === 0) continue
    rules.push(
      `.page__content [data-char-style="${attributeText(style.id)}"] { ${declarations(css).join(' ')} }`,
    )
  }

  for (const style of Object.values(sheet.styles)) {
    if (style.type !== StyleType.Paragraph) continue
    rules.push(
      rule(`.page__content > [data-style-id="${attributeText(style.id)}"]`, resolveStyle(sheet, style.id)),
    )
  }

  // O sumário é um bloco no editor e parágrafos com `toc 1`, `toc 2`… no arquivo.
  // Repetidas no fim, e na mesma ordem, para valer a mesma precedência.
  const inContents = rules
    .filter((text) => text.startsWith('.page__content > ') && !text.startsWith('.page__content > * + *'))
    .map((text) => text.replace('.page__content > ', '.page__content > [data-toc] > '))

  return [...rules, ...inContents].join('\n')
}

function rule(selector: string, style: ResolvedStyle): string {
  return `${selector} { ${declarations([...paragraphCss(style), ...characterCss(style)]).join(' ')} }`
}

function paragraphCss({ paragraph }: ResolvedStyle): Array<[string, string]> {
  const indent = millimeters(paragraph.indentMm)
  const indentRight = millimeters(paragraph.indentRightMm)
  return [
    ['margin-top', points(paragraph.spaceBefore)],
    ['margin-bottom', points(paragraph.spaceAfter)],
    ['text-align', paragraph.textAlign ?? 'start'],
    ['padding-left', indent],
    ['--recuo', indent],
    ['padding-right', indentRight],
    ['--recuo-direita', indentRight],
    ['text-indent', millimeters(paragraph.firstLineMm)],
    ['background-color', paragraph.background ?? 'transparent'],
  ]
}

function characterCss({ paragraph, character }: ResolvedStyle): Array<[string, string]> {
  const family = character.fontFamily ?? null
  const decorations = [
    character.underline === true ? 'underline' : '',
    character.strike === true ? 'line-through' : '',
  ]
    .filter((line) => line.length > 0)
    .join(' ')

  const css: Array<[string, string]> = [
    ['font-size', character.fontSize ?? '12pt'],
    ['font-weight', character.bold === true ? '700' : '400'],
    ['font-style', character.italic === true ? 'italic' : 'normal'],
    ['text-decoration', decorations.length > 0 ? decorations : 'none'],
    ['text-transform', character.allCaps === true ? 'uppercase' : 'none'],
    ['font-variant', character.smallCaps === true ? 'small-caps' : 'normal'],
    ['line-height', lineHeightOf(paragraph, family)],
  ]
  if (family !== null) css.unshift(['font-family', fontStackOf(family)])
  // Sem cor declarada vale o "automático" do Word, que não é cor a gravar.
  if (character.color !== undefined) css.push(['color', character.color])
  return css
}

/** O que o estilo de caractere cala é do parágrafo. */
function declaredCharacterCss(character: StyleCharacterFormat): Array<[string, string]> {
  const css: Array<[string, string]> = []
  if (character.fontFamily !== undefined) css.push(['font-family', fontStackOf(character.fontFamily)])
  if (character.fontSize !== undefined) css.push(['font-size', character.fontSize])
  if (character.bold !== undefined) css.push(['font-weight', character.bold ? '700' : '400'])
  if (character.italic !== undefined) css.push(['font-style', character.italic ? 'italic' : 'normal'])
  if (character.underline === true || character.strike === true) {
    const lines = [
      character.underline === true ? 'underline' : '',
      character.strike === true ? 'line-through' : '',
    ]
    css.push(['text-decoration', lines.filter((line) => line !== '').join(' ')])
  }
  if (character.allCaps !== undefined) css.push(['text-transform', character.allCaps ? 'uppercase' : 'none'])
  if (character.smallCaps !== undefined)
    css.push(['font-variant', character.smallCaps ? 'small-caps' : 'normal'])
  if (character.color !== undefined) css.push(['color', character.color])
  if (character.highlight !== undefined) css.push(['background-color', character.highlight])
  return css
}

function lineHeightOf(paragraph: ResolvedStyle['paragraph'], family: string | null): string {
  const spacing = paragraph.lineSpacing
  if (spacing === undefined) return cssLineHeightOf(1, family)
  switch (spacing.kind) {
    case 'multiple':
      return cssLineHeightOf(usableFactor(spacing.factor), family)
    case 'exact':
      return `${spacing.pt}pt`
    case 'atLeast':
      return `max(${spacing.pt}pt, ${explicitCssLineHeightOf(1, family)}em)`
  }
}

/** O nome solto das tabelas embutidas ganha aspas e uma família genérica. */
function fontStackOf(family: string): string {
  if (family.includes(',')) return family
  const generic = /serif|times|roman|cambria|georgia|garamond/i.test(family) ? 'serif' : 'sans-serif'
  return `'${family.replace(/'/g, '')}', ${generic}`
}

function declarations(pairs: ReadonlyArray<readonly [string, string]>): string[] {
  return pairs.map(([property, value]) => `${property}: ${value};`)
}

function points(value: number | undefined): string {
  return `${value ?? 0}pt`
}

function millimeters(value: number | undefined): string {
  return `${value ?? 0}mm`
}

/** O id vai entre aspas num seletor de atributo: só aspas e barra precisam de escape. */
function attributeText(id: string): string {
  return id.replace(/["\\]/g, '\\$&')
}
