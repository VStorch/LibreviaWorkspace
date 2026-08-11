/**
 * Busca de texto.
 *
 * Fica na camada pura para poder ser testada sem editor: é a parte da
 * pesquisa que mais fácil erra em detalhe (sobreposição, acentuação,
 * maiúsculas) e a que menos precisa do ProseMirror para ser verificada.
 */

export interface Occurrence {
  readonly start: number
  readonly end: number
}

/**
 * Ocorrências não sobrepostas de `needle` em `haystack`.
 *
 * A busca avança pelo comprimento do termo encontrado — procurar "aa" em
 * "aaaa" devolve duas ocorrências, não três. É o comportamento que faz
 * "substituir tudo" terminar, em vez de reprocessar o que acabou de escrever.
 */
export function findOccurrences(haystack: string, needle: string, caseSensitive = false): Occurrence[] {
  if (needle.length === 0) return []

  const subject = caseSensitive ? haystack : haystack.toLowerCase()
  const target = caseSensitive ? needle : needle.toLowerCase()

  const found: Occurrence[] = []
  let index = subject.indexOf(target)

  while (index !== -1) {
    found.push({ start: index, end: index + needle.length })
    index = subject.indexOf(target, index + target.length)
  }

  return found
}

/** Índice da próxima ocorrência ao navegar de forma circular. */
export function stepIndex(current: number, total: number, delta: number): number {
  if (total === 0) return -1
  if (current < 0) return delta > 0 ? 0 : total - 1
  return (current + delta + total) % total
}
