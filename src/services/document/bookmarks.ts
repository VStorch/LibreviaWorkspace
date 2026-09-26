/**
 * As regras dos marcadores (bookmarks) do Word, sem editor.
 *
 * No arquivo um marcador é um par `w:bookmarkStart`/`w:bookmarkEnd` ligado por um
 * `w:id`, com um nome no início. O nome é o que o sumário, o link interno e a
 * referência cruzada citam; o id só casa as duas pontas, e tem de ser único.
 */

/** O maior nome que o Word aceita. */
export const BOOKMARK_NAME_MAX = 40

/**
 * O nome que o Word deixa a pessoa dar: começa por letra e segue com letras,
 * algarismos e sublinhado, até 40 caracteres. Sem espaço — o Word recusa, e o
 * campo que o cita (`REF nome`) partiria nele.
 *
 * O sublinhado inicial fica de fora de propósito: é a marca dos ocultos que o
 * próprio Word cria (`_Toc…`, `_Ref…`), e um nome assim dado à mão sumiria da
 * lista com os outros.
 */
export function isValidBookmarkName(name: string): boolean {
  return new RegExp(`^\\p{L}[\\p{L}\\p{N}_]{0,${BOOKMARK_NAME_MAX - 1}}$`, 'u').test(name)
}

/** Oculto: os que o Word cria para o sumário e para as referências. */
export function isHiddenBookmark(name: string): boolean {
  return name.startsWith('_')
}

/**
 * O id do marcador novo: um a mais que o maior do documento.
 *
 * Numérico porque o `w:id` é um inteiro no esquema; o que vier de fora e não for
 * número não entra na conta, mas também não colide — é outro texto.
 */
export function nextBookmarkId(existing: Iterable<string>): string {
  let highest = -1
  for (const id of existing) {
    const value = Number(id)
    if (Number.isInteger(value) && value > highest) highest = value
  }
  return String(highest + 1)
}

/**
 * Um nome oculto novo com o prefixo dado — `_Ref` para a referência cruzada,
 * `_Toc` para o sumário.
 *
 * O Word sorteia os algarismos; aqui eles contam a partir do maior que já existe,
 * que dá o mesmo formato sem depender de sorte para não repetir.
 */
export function hiddenBookmarkName(prefix: '_Ref' | '_Toc', existing: Iterable<string>): string {
  let highest = 0
  const pattern = new RegExp(`^${prefix}(\\d+)$`)
  for (const name of existing) {
    const match = pattern.exec(name)
    if (match !== null) highest = Math.max(highest, Number(match[1]))
  }
  return `${prefix}${String(highest + 1).padStart(9, '0')}`
}
