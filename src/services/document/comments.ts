import type { DocumentComment, DocumentNode } from './model.js'

/**
 * Como as seções (`resolveSections`): a biblioteca só ganha entradas, e o texto
 * diz quais valem, para o desfazer tirar e devolver o comentário inteiro.
 * `outside` são as conversas ancoradas fora do corpo, que valem sem pontas.
 */
export function resolveComments(
  anchored: ReadonlySet<string>,
  library: readonly DocumentComment[],
  outside: ReadonlySet<string> = new Set(),
): DocumentComment[] {
  const byId = new Map(library.map((comment) => [comment.id, comment]))
  const exists = new Map<string, boolean>()
  const check = (comment: DocumentComment, depth: number): boolean => {
    const known = exists.get(comment.id)
    if (known !== undefined) return known
    const parent = comment.parentId === undefined ? undefined : byId.get(comment.parentId)
    // Resposta sem o comentário na biblioteca vira conversa própria; o teto de
    // profundidade só segura o ciclo.
    const result =
      parent === undefined || depth > 100
        ? anchored.has(comment.id) || outside.has(comment.id)
        : check(parent, depth + 1)
    exists.set(comment.id, result)
    return result
  }
  return library.filter((comment) => check(comment, 0))
}

/** A conversa a que o comentário pertence: o id do que a abre. */
export function threadRootOf(library: readonly DocumentComment[], id: string): string {
  const byId = new Map(library.map((comment) => [comment.id, comment]))
  let current = byId.get(id)
  for (let depth = 0; current?.parentId !== undefined && depth < 100; depth++) {
    const parent = byId.get(current.parentId)
    if (parent === undefined) break
    current = parent
  }
  return current?.id ?? id
}

export function commentAnchorIdsOfJson(doc: DocumentNode): Set<string> {
  const ids = new Set<string>()
  const walk = (node: DocumentNode): void => {
    if (node.type === 'commentStart' || node.type === 'commentEnd') {
      ids.add(String(node.attrs?.['cid'] ?? ''))
      return
    }
    for (const child of node.content ?? []) walk(child)
  }
  walk(doc)
  return ids
}

/** As que o arquivo ancora fora do corpo. */
export function commentsOutsideOf(doc: DocumentNode, library: readonly DocumentComment[]): readonly string[] {
  const anchored = commentAnchorIdsOfJson(doc)
  return library
    .filter((comment) => threadRootOf(library, comment.id) === comment.id && !anchored.has(comment.id))
    .map((comment) => comment.id)
}

/** Da biblioteca inteira: o comentário desfeito volta com o refazer, e o id dele não pode ter ido a outro. */
export function nextCommentId(library: readonly DocumentComment[]): string {
  let max = -1
  for (const comment of library) {
    if (/^\d+$/.test(comment.id)) max = Math.max(max, Number(comment.id))
  }
  return String(max + 1)
}

/** Como o Word: a primeira letra de cada palavra, até três. */
export function initialsOf(author: string): string {
  return author
    .split(/\s+/)
    .filter((word) => word !== '')
    .slice(0, 3)
    .map((word) => word[0]!.toLocaleUpperCase())
    .join('')
}

export function paragraphsOfText(text: string): string[] {
  return text.replace(/\r\n?/g, '\n').split('\n')
}
