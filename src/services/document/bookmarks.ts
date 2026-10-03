/** O nome é o que o sumário e as referências citam; o `w:id` casa as duas pontas e é único. */

export const BOOKMARK_NAME_MAX = 40

/**
 * Como o Word: letra, depois letras, algarismos e sublinhado, até 40. Sem o
 * sublinhado inicial, que marca os ocultos (`_Toc…`, `_Ref…`).
 */
export function isValidBookmarkName(name: string): boolean {
  return new RegExp(`^\\p{L}[\\p{L}\\p{N}_]{0,${BOOKMARK_NAME_MAX - 1}}$`, 'u').test(name)
}

/** Oculto: os que o Word cria para o sumário e para as referências. */
export function isHiddenBookmark(name: string): boolean {
  return name.startsWith('_')
}

/** O `w:id` é inteiro no esquema; o que não for número não entra na conta. */
export function nextBookmarkId(existing: Iterable<string>): string {
  let highest = -1
  for (const id of existing) {
    const value = Number(id)
    if (Number.isInteger(value) && value > highest) highest = value
  }
  return String(highest + 1)
}

/** O Word sorteia os algarismos; aqui contam a partir do maior, sem depender de sorte. */
export function hiddenBookmarkName(prefix: '_Ref' | '_Toc', existing: Iterable<string>): string {
  let highest = 0
  const pattern = new RegExp(`^${prefix}(\\d+)$`)
  for (const name of existing) {
    const match = pattern.exec(name)
    if (match !== null) highest = Math.max(highest, Number(match[1]))
  }
  return `${prefix}${String(highest + 1).padStart(9, '0')}`
}
