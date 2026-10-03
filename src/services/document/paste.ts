import type { DocumentNode } from './model.js'

/**
 * Uma linha só devolve texto puro, para não partir a frase onde se cola. Várias
 * linhas viram um parágrafo por linha, inclusive os vazios.
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
 * `\v` e `\f` viram quebra, porque é assim que uma célula do Excel e um PDF
 * antigo separam linhas; os outros caracteres de controle saem, e a última
 * quebra de linha também.
 */
function normalize(raw: string): string {
  const lines = raw.replace(/\r\n?/gu, '\n').replace(/[\v\f\u0085\u2028\u2029]/gu, '\n')

  // Caractere por caractere, e não por expressão regular com `\u0000-\u001f`,
  // que a regra `no-control-regex` proíbe.
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
