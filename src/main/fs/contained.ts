import { join, resolve } from 'node:path'

/**
 * Checked against the resolved path, not the text: the name comes from a URL any page can write,
 * and `..` or an absolute name would leave the folder.
 */
export function containedPath(root: string, name: string): string | null {
  if (name === '') return null
  const inside = resolve(root)
  const candidate = resolve(inside, name)
  if (candidate !== join(inside, name) || !candidate.startsWith(inside)) return null
  return candidate
}
