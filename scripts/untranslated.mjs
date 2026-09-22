/**
 * Encontra frase em portugues ainda presa no codigo.
 *
 * A traducao da interface e uma varredura grande e mecanica, e o compilador so
 * cobre metade dela: ele garante que toda chave existe nos dois idiomas, e nao
 * tem como saber que a frase nunca virou chave. Este script e a outra metade -
 * ele conta o que falta, e por isso o progresso da varredura e um numero em vez
 * de uma impressao.
 *
 *   node scripts/untranslated.mjs          lista tudo, sai 1 se houver algo
 *   node scripts/untranslated.mjs --count  so o numero
 *
 * ## O que ele olha, e o que nao
 *
 * Comentario e removido antes de qualquer coisa: o repositorio comenta em
 * portugues de proposito, e isso nao muda. `*.test.ts` e `e2e/` ficam de fora
 * pelo mesmo motivo - descricao de teste e portugues por regra.
 *
 * O que sobra e string com letra acentuada ou com palavra de funcao portuguesa,
 * que e um sinal forte de frase escrita para uma pessoa ler. Da falso positivo
 * (um nome de fonte, uma chave de formato) e por isso existe a lista de
 * excecoes no fim - cada uma com o motivo escrito.
 */
import { readdirSync, readFileSync, statSync } from 'node:fs'
import { join, relative, sep } from 'node:path'

const ROOT = process.cwd()
const SRC = join(ROOT, 'src')

/** Pastas que a varredura nao entra. */
const SKIP_DIRS = new Set(['node_modules', 'out', 'dist', 'release', '.git'])

/**
 * Arquivos isentos, com o motivo.
 *
 * Isencao e divida, nao permissao: cada linha aqui e uma frase que a pessoa que
 * escolheu ingles ainda vai ler em portugues.
 */
const EXEMPT = new Map([
  // O catalogo E o portugues. Procurar frases em portugues nele acharia todas.
  ['src/shared/i18n', 'o catalogo guarda as duas linguas de proposito'],
  // Nomes de funcao de planilha (SOMA, PROCV). Sao sintaxe que o arquivo .xlsx
  // guarda, e nao texto de interface - traduzi-los quebraria as formulas.
  ['src/services/spreadsheet/formula', 'nome de funcao e sintaxe, nao interface'],
  // O campo `does:` da tabela de atalhos e documentacao: nenhum codigo o
  // renderiza. Confirmado por busca - nada le `.does`. Se um dia alguem
  // mostrar a tabela de atalhos na tela, tire daqui antes.
  ['src/shared/shortcuts.ts', 'o campo does e documentacao, nao vai para a tela'],
])

/** Palavras que denunciam uma frase portuguesa mesmo sem acento. */
const PT_WORDS =
  /\b(de|da|do|das|dos|para|com|sem|nao|nenhum|nenhuma|salvar|abrir|fechar|arquivo|pagina|linha|coluna|tabela|imagem|texto|erro|aviso)\b/i

function walk(dir, out = []) {
  for (const name of readdirSync(dir)) {
    if (SKIP_DIRS.has(name)) continue
    const path = join(dir, name)
    if (statSync(path).isDirectory()) walk(path, out)
    else if (/\.tsx?$/.test(path) && !/\.test\.tsx?$/.test(path)) out.push(path)
  }
  return out
}

/** Tira comentario de bloco e de linha, preservando a contagem de linhas. */
function stripComments(source) {
  return source
    .replace(/\/\*[\s\S]*?\*\//g, (block) => block.replace(/[^\n]/g, ' '))
    .split('\n')
    .map((line) => line.replace(/(^|[^:])\/\/.*$/, '$1'))
    .join('\n')
}

function isPortuguese(text) {
  if (text.trim().length < 3) return false
  if (!/[A-Za-zÀ-ÿ]{3}/.test(text)) return false
  return /[À-ÿ]/.test(text) || PT_WORDS.test(text)
}

function exemptionFor(relPath) {
  const posix = relPath.split(sep).join('/')
  for (const [prefix, reason] of EXEMPT) {
    if (posix.startsWith(prefix)) return reason
  }
  return null
}

const findings = []

for (const file of walk(SRC)) {
  const relPath = relative(ROOT, file)
  if (exemptionFor(relPath) !== null) continue

  const source = stripComments(readFileSync(file, 'utf8'))
  const lines = source.split('\n')

  lines.forEach((line, index) => {
    // Uma linha que ja chama o tradutor nao esta pendente, mesmo que traga
    // portugues junto - o argumento de `t()` e a chave, nao a frase.
    if (/\bt\(\s*['"]/.test(line)) return

    const quoted = [...line.matchAll(/(['"`])((?:\\.|(?!\1)[^\\\n])*)\1/g)].map((m) => m[2])
    const jsxText = [...line.matchAll(/>([^<>{}]+)</g)].map((m) => m[1])

    for (const text of [...quoted, ...jsxText]) {
      if (!isPortuguese(text)) continue
      findings.push({ file: relPath.split(sep).join('/'), line: index + 1, text: text.trim() })
    }
  })
}

if (process.argv.includes('--count')) {
  console.log(findings.length)
} else {
  for (const found of findings) {
    console.log(`${found.file}:${found.line}: ${found.text}`)
  }
  console.log(`\n${findings.length} frase(s) ainda no codigo.`)
}

process.exit(findings.length === 0 ? 0 : 1)
