import { readdirSync, readFileSync, statSync } from 'node:fs'
import { join, relative, sep } from 'node:path'
import { describe, expect, it } from 'vitest'

const SRC = join(process.cwd(), 'src')
const AI_FOLDER = 'src/main/ai/'

/** Each exception with what it reaches; the AI folder is the only one that goes out. */
const ALLOWED: Readonly<Record<string, readonly string[]>> = {
  'src/main/pdf-scheme.ts': ['electron net'],
  'src/renderer/pdf-viewer/ranged-source.ts': ['fetch'],
  'src/main/sidecar/client.ts': ['child_process'],
  'src/main/system-fonts.ts': ['child_process'],
}

const PATTERNS: readonly (readonly [string, RegExp])[] = [
  ['fetch', /(?<![\w.$])fetch\s*\(/u],
  ['WebSocket', /\bnew\s+WebSocket\b/u],
  ['EventSource', /\bEventSource\b/u],
  ['XMLHttpRequest', /\bXMLHttpRequest\b/u],
  ['node network module', /from\s+['"](node:)?(http|https|http2|net|tls|dgram)['"]/u],
  ['undici', /from\s+['"]undici['"]/u],
  ['@anthropic-ai/sdk', /from\s+['"]@anthropic-ai\/sdk['"]/u],
  ['@agentclientprotocol/sdk', /from\s+['"]@agentclientprotocol\/sdk[/'"]/u],
  ['electron net', /import\s*\{[^}]*\bnet\b[^}]*\}\s*from\s+['"]electron['"]/u],
  ['child_process', /['"](node:)?child_process['"]/u],
]

function sourceFiles(dir: string, out: string[] = []): string[] {
  for (const name of readdirSync(dir)) {
    const path = join(dir, name)
    if (statSync(path).isDirectory()) sourceFiles(path, out)
    else if (/\.tsx?$/u.test(name) && !/\.test\.tsx?$/u.test(name)) out.push(path)
  }
  return out
}

describe('network boundary', () => {
  it('only src/main/ai reaches the network or spawns processes, outside the known exceptions', () => {
    const findings = sourceFiles(SRC).flatMap((path) => {
      const file = relative(process.cwd(), path).split(sep).join('/')
      if (file.startsWith(AI_FOLDER)) return []
      const source = readFileSync(path, 'utf8')
      return PATTERNS.filter(
        ([name, pattern]) => pattern.test(source) && !(ALLOWED[file] ?? []).includes(name),
      ).map(([name]) => `${file}: ${name}`)
    })
    expect(findings).toEqual([])
  })

  it('the exceptions still hold; one that disappeared leaves the list', () => {
    const stale = Object.entries(ALLOWED).flatMap(([file, names]) => {
      const source = readFileSync(join(process.cwd(), file), 'utf8')
      return names
        .filter((name) => !PATTERNS.find(([key]) => key === name)![1].test(source))
        .map((name) => `${file}: ${name}`)
    })
    expect(stale).toEqual([])
  })
})
