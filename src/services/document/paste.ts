import type { DocumentNode } from './model.js'

/**
 * Colar sem formatação, em forma de conteúdo do documento.
 *
 * O texto vem da área de transferência do sistema, que só o processo main
 * alcança; aqui ele é traduzido para nós do documento. Passa por uma função pura
 * para que o caso difícil — quebra de linha do Windows, tabulação, caractere de
 * controle que veio de um PDF — seja verificável sem abrir um editor.
 */

/**
 * O texto colado, já como conteúdo para inserir no cursor.
 *
 * Uma linha só devolve texto puro: inserir um parágrafo no meio de uma frase
 * partiria a frase em duas, e colar o nome de uma cidade dentro de um endereço
 * é o caso comum.
 *
 * Várias linhas devolvem um parágrafo por linha — inclusive os vazios, porque
 * quem copiou um trecho com linha em branco espera vê-la de volta.
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
 * Normaliza o que a área de transferência entrega.
 *
 * - quebra de linha do Windows e do Mac clássico viram `\n`;
 * - `\v` e `\f` também: é assim que uma célula do Excel e um PDF antigo separam
 *   linhas, e como caractere de controle eles apareceriam como quadradinho;
 * - o resto dos caracteres de controle sai fora — um `\u0000` colado de um
 *   arquivo binário deixaria o documento impossível de salvar;
 * - a última quebra de linha é descartada: texto copiado de terminal, de célula
 *   de planilha ou de tabela vem com ela, e ninguém quer o parágrafo vazio a
 *   mais no fim.
 */
function normalize(raw: string): string {
  const lines = raw.replace(/\r\n?/gu, '\n').replace(/[\v\f\u0085\u2028\u2029]/gu, '\n')

  // Caractere por caractere, e não por expressão regular: uma classe com
  // `\u0000-\u001f` dentro é justamente o que a regra `no-control-regex` proíbe,
  // e com razão — ela esconde o que está sendo apagado atrás de uma faixa de
  // números. Aqui está escrito: controle nenhum passa, menos os dois que são
  // conteúdo.
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
