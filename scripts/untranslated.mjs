/**
 * Finds Portuguese sentences stuck in code, outside the catalog. The compiler guarantees every key
 * exists in both languages, and knows nothing of a sentence that never became a key.
 *
 *   node scripts/untranslated.mjs          lists everything, exits 1 if anything is found
 *   node scripts/untranslated.mjs --count  only the number
 *
 * Comments, `*.test.ts` and `e2e/` are skipped: comments and tests are in English, by rule, and
 * tests hold Portuguese only as test data. The rest is strings with an accented letter or a
 * Portuguese function word; false positives go to the exemption list, with the reason.
 */
import { readdirSync, readFileSync, statSync } from 'node:fs'
import { join, relative, sep } from 'node:path'

const ROOT = process.cwd()
const SRC = join(ROOT, 'src')

/** Folders the scan does not enter. */
const SKIP_DIRS = new Set(['node_modules', 'out', 'dist', 'release', '.git'])

/**
 * Exempt files, with the reason. An exemption is a sentence that someone who chose English reads in
 * Portuguese.
 */
const EXEMPT = new Map([
  ['src/shared/i18n', 'the catalog holds both languages on purpose'],
  // SOMA and PROCV are syntax the .xlsx stores: translating them would break the formulas.
  ['src/services/spreadsheet/formula', 'function names and syntax, not interface'],
  // No code renders `.does`. If the shortcut table ever reaches the screen, remove this.
  ['src/shared/shortcuts.ts', 'the does field is documentation, it never reaches the screen'],
  ['src/main/context-menu.ts', 'technical IPC log in the main terminal, never on screen'],
  ['src/main/sidecar/index.ts', 'technical sidecar startup log in the main terminal, never on screen'],
  ['src/main/spellcheck.ts', 'technical spell checker logs in the main terminal, never on screen'],
])

/** Words that give away a Portuguese sentence even without accents. */
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

/** Strips block and line comments, keeping the line count. */
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
    // A line already calling `t()` is not pending: the argument is the key.
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
