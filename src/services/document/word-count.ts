import type { DocumentNode } from './model.js'

/** Palavras e caracteres vêm do `CharacterCount`, como na barra de status, para não haver dois números. */

/** Todo espaço do Unicode, como o `\u00a0` que o Word põe em datas e números. */
export function charactersWithoutSpaces(text: string): number {
  return [...text.replace(/\s|\u00a0/gu, '')].length
}

/** Lista, citação e célula **contêm** parágrafos: contá-los somaria cada linha duas vezes. */
const PARAGRAPH_TYPES = new Set(['paragraph', 'heading', 'codeBlock'])

/** Como o Word, só linha com texto; a marca de seção, que é parágrafo vazio, também fica de fora. */
export function countParagraphs(node: DocumentNode): number {
  let total = 0

  const visit = (current: DocumentNode): void => {
    if (PARAGRAPH_TYPES.has(current.type)) {
      if (inlineTextOf(current).trim() !== '') total += 1
      return
    }

    for (const child of current.content ?? []) visit(child)
  }

  visit(node)
  return total
}

function inlineTextOf(node: DocumentNode): string {
  if (typeof node.text === 'string') return node.text
  return (node.content ?? []).map(inlineTextOf).join('')
}
