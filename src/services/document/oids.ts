/** At any depth: a DOCX list item carries its own. */
export function withoutOids(node: unknown): unknown {
  if (Array.isArray(node)) return node.map(withoutOids)
  if (typeof node !== 'object' || node === null) return node
  const record = node as Record<string, unknown>
  const { attrs, content } = record
  const copy: Record<string, unknown> = { ...record }
  if (typeof attrs === 'object' && attrs !== null && 'oid' in attrs) {
    const rest = Object.entries(attrs).filter(([name]) => name !== 'oid')
    if (rest.length > 0) copy['attrs'] = Object.fromEntries(rest)
    else delete copy['attrs']
  }
  if (Array.isArray(content)) copy['content'] = content.map(withoutOids)
  return copy
}
