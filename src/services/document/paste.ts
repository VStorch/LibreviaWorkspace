import type { DocumentNode } from './model.js'

/**
 * A single line returns plain text, so the sentence is not split where it is pasted. Several lines
 * become one paragraph per line, empty ones included.
 */
export function plainPasteContent(raw: string): DocumentNode[] {
  const text = normalize(raw)
  if (text === '') return []

  if (!text.includes('\n')) return [{ type: 'text', text }]

  return text
    .split('\n')
    .map<DocumentNode>((line) =>
      line === '' ? { type: 'paragraph' } : { type: 'paragraph', content: [{ type: 'text', text: line }] },
    )
}

/**
 * `\v` and `\f` become breaks, because that is how an Excel cell and an old PDF separate lines;
 * other control characters are dropped, and so is the final line break.
 */
function normalize(raw: string): string {
  const lines = raw.replace(/\r\n?/gu, '\n').replace(/[\v\f\u0085\u2028\u2029]/gu, '\n')

  // Character by character, not a regular expression with `\u0000-\u001f`, which the
  // `no-control-regex` rule forbids.
  return [...lines]
    .filter((char) => !isControl(char))
    .join('')
    .replace(/\n$/u, '')
}

function isControl(char: string): boolean {
  const code = char.codePointAt(0) ?? 0
  if (char === '\n' || char === '\t') return false
  return code < 0x20 || code === 0x7f
}
