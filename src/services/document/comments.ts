import type { DocumentComment, DocumentNode } from './model.js'

/**
 * Quais comentários o documento tem agora (M10, fase 2).
 *
 * Mesmo desenho das seções (ver `resolveSections`): a loja guarda uma
 * **biblioteca** que só ganha entradas, e quem diz quais valem é o texto. A
 * conversa existe enquanto as pontas dela estão no documento; a resposta, enquanto
 * a conversa que ela responde existe. Assim o desfazer do editor — que só conhece
 * o texto — tira e devolve o comentário inteiro, cartão incluído, sem a loja
 * entrar no histórico.
 *
 * `outside` são as conversas que o arquivo ancora fora do corpo (cabeçalho, nota,
 * caixa de texto): o editor não tem as pontas delas, e elas valem assim mesmo.
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
    // A resposta cujo comentário não está na biblioteca vira conversa própria,
    // como o painel a mostra; o teto de profundidade só segura o ciclo.
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

/** As pontas de comentário de um documento em JSON — o `cid` de cada uma. */
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

/**
 * As conversas sem ponta no texto quando o documento abriu: as que o arquivo
 * ancora fora do corpo. Ver `resolveComments`.
 */
export function commentsOutsideOf(doc: DocumentNode, library: readonly DocumentComment[]): readonly string[] {
  const anchored = commentAnchorIdsOfJson(doc)
  return library
    .filter((comment) => threadRootOf(library, comment.id) === comment.id && !anchored.has(comment.id))
    .map((comment) => comment.id)
}

/**
 * O id do comentário novo: o próximo número depois do maior da biblioteca.
 *
 * Da biblioteca, e não só dos que valem agora: o comentário desfeito volta com o
 * refazer, e o id dele não pode ter sido dado a outro no meio.
 */
export function nextCommentId(library: readonly DocumentComment[]): string {
  let max = -1
  for (const comment of library) {
    if (/^\d+$/.test(comment.id)) max = Math.max(max, Number(comment.id))
  }
  return String(max + 1)
}

/** As iniciais do nome, como o Word as tira: a primeira letra de cada palavra, até três. */
export function initialsOf(author: string): string {
  return author
    .split(/\s+/)
    .filter((word) => word !== '')
    .slice(0, 3)
    .map((word) => word[0]!.toLocaleUpperCase())
    .join('')
}

/** O texto da caixa do painel em parágrafos do comentário — uma linha, um parágrafo. */
export function paragraphsOfText(text: string): string[] {
  return text.replace(/\r\n?/g, '\n').split('\n')
}
