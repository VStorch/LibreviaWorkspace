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

  const subject = caseSensitive ? haystack : foldCase(haystack)
  const target = caseSensitive ? needle : foldCase(needle)

  const found: Occurrence[] = []
  let index = subject.indexOf(target)

  while (index !== -1) {
    found.push({ start: index, end: index + needle.length })
    index = subject.indexOf(target, index + target.length)
  }

  return found
}

/**
 * Minúsculas **sem mudar o comprimento** do texto.
 *
 * `toLowerCase()` não preserva comprimento: o `İ` turco (U+0130) vira duas
 * unidades, `i` mais um ponto combinante. As posições que esta busca devolve são
 * as posições do documento — é com elas que o destaque é desenhado e o
 * "substituir tudo" corta o texto —, e um caractere a mais no meio do caminho
 * desloca tudo o que vem depois: o realce cai sobre a palavra errada e a
 * substituição come uma letra do vizinho.
 *
 * Então cada ponto de código é convertido sozinho, e só quando a conversão couber
 * no mesmo espaço. O que não couber fica como está — deixa de casar sem acento,
 * que é o mesmo acordo já feito com a acentuação, e é muito menos grave do que
 * um índice errado.
 */
function foldCase(text: string): string {
  // Ponto de código, e não unidade de UTF-16: um par substituto convertido pela
  // metade produziria lixo, e as duas unidades contam igual na soma.
  let folded = ''
  for (const character of text) {
    const lower = character.toLowerCase()
    folded += lower.length === character.length ? lower : character
  }

  return folded
}

/** Índice da próxima ocorrência ao navegar de forma circular. */
export function stepIndex(current: number, total: number, delta: number): number {
  if (total === 0) return -1
  if (current < 0) return delta > 0 ? 0 : total - 1
  return (current + delta + total) % total
}
