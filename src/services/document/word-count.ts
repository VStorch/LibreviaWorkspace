import type { DocumentNode } from './model.js'

/**
 * As contagens que o diálogo "Contar palavras" mostra e que a extensão
 * `CharacterCount` **não** dá: caracteres sem espaço e parágrafos.
 *
 * Palavras e caracteres continuam vindo do `CharacterCount`, que já alimenta a
 * barra de status. Recontá-los aqui de outro jeito produziria dois números para
 * a mesma coisa na mesma tela — e o usuário ficaria sem saber em qual acreditar.
 */

/**
 * Caracteres sem contar espaço nenhum.
 *
 * Toda categoria de espaço do Unicode, e não só o `' '`: um documento vindo do
 * Word traz espaço inquebrável (`\u00a0`) em data, em número e antes de unidade,
 * e contá-lo como caractere de texto daria um total maior que o do Word sem
 * razão visível na tela.
 */
export function charactersWithoutSpaces(text: string): number {
  return [...text.replace(/\s|\u00a0/gu, '')].length
}

/**
 * Nós que são uma linha de texto do documento.
 *
 * `blockquote`, `listItem` e célula de tabela ficam de fora de propósito: eles
 * **contêm** parágrafos, e contar os dois lados somaria cada linha duas vezes.
 */
const PARAGRAPH_TYPES = new Set(['paragraph', 'heading', 'codeBlock'])

/**
 * Parágrafos, na conta do Word: linha com texto.
 *
 * Parágrafo vazio não entra — é tecla Enter batida para abrir espaço, não
 * conteúdo, e o Word também não o conta. A marca de seção (`w:sectPr` guardado
 * num parágrafo vazio) é descartada pela mesma regra, o que é conveniente: ela
 * não é um parágrafo do texto, é estrutura do arquivo.
 */
export function countParagraphs(node: DocumentNode): number {
  let total = 0

  const visit = (current: DocumentNode): void => {
    if (PARAGRAPH_TYPES.has(current.type)) {
      if (inlineTextOf(current).trim() !== '') total += 1
      // Um parágrafo não tem parágrafo dentro: descer daqui só encontraria
      // marcas de texto.
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
