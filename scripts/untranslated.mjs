/**
 * Encontra frase em portugues presa no codigo, fora do catalogo. O compilador
 * garante que toda chave existe nos dois idiomas, e nao sabe da frase que nunca
 * virou chave.
 *
 *   node scripts/untranslated.mjs          lista tudo, sai 1 se houver algo
 *   node scripts/untranslated.mjs --count  so o numero
 *
 * Comentarios, `*.test.ts` e `e2e/` ficam de fora: sao em portugues por regra. O
 * resto e string com letra acentuada ou palavra de funcao portuguesa; os falsos
 * positivos vao para a lista de excecoes, com o motivo.
 */
import { readdirSync, readFileSync, statSync } from 'node:fs'
import { join, relative, sep } from 'node:path'

const ROOT = process.cwd()
const SRC = join(ROOT, 'src')

/** Pastas que a varredura nao entra. */
const SKIP_DIRS = new Set(['node_modules', 'out', 'dist', 'release', '.git'])

/** Arquivos isentos, com o motivo. Isencao e frase que quem escolheu ingles le em portugues. */
const EXEMPT = new Map([
  ['src/shared/i18n', 'o catalogo guarda as duas linguas de proposito'],
  // SOMA e PROCV sao sintaxe que o .xlsx guarda: traduzi-los quebraria as formulas.
  ['src/services/spreadsheet/formula', 'nome de funcao e sintaxe, nao interface'],
  // Nenhum codigo renderiza `.does`. Se a tabela de atalhos for para a tela, tire daqui.
  ['src/shared/shortcuts.ts', 'o campo does e documentacao, nao vai para a tela'],
  ['src/main/context-menu.ts', 'log tecnico de IPC no terminal do main, nao vai para a tela'],
  ['src/main/sidecar/index.ts', 'log tecnico de subida do sidecar no terminal do main, nao vai para a tela'],
  ['src/main/spellcheck.ts', 'logs tecnicos do corretor no terminal do main, nao vao para a tela'],
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
    // A linha que ja chama `t()` nao esta pendente: o argumento e a chave.
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
