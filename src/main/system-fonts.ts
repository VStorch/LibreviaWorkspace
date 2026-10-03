import { execFile } from 'node:child_process'
import { join } from 'node:path'
import { promisify } from 'node:util'
import { parseFontconfigFamilies, parseWindowsFontRegistry } from '@services/document/font-list.js'

/**
 * `queryLocalFonts()` pede permissão e não vale num renderer sem origem: quem
 * pergunta é o main. Conforto, e não requisito: toda falha devolve lista vazia.
 */

const run = promisify(execFile)

/** Para o caso patológico, como o cache do fontconfig sendo reconstruído. */
const TIMEOUT_MS = 4000
const MAX_OUTPUT_BYTES = 8 * 1024 * 1024

/** Uma vez por sessão: instalar fonte com o aplicativo aberto é raro. */
let cached: Promise<string[]> | null = null

export function listInstalledFontFamilies(): Promise<string[]> {
  cached ??= collect().then((families) => {
    // Lista vazia não fica em cache: pode ser falha transitória, e guardá-la
    // deixaria a barra sem fonte nenhuma pela sessão inteira.
    if (families.length === 0) cached = null
    return families
  })
  return cached
}

/** Para os testes. */
export function forgetInstalledFonts(): void {
  cached = null
}

async function collect(): Promise<string[]> {
  try {
    return process.platform === 'win32' ? await fromWindowsRegistry() : await fromFontconfig()
  } catch {
    return []
  }
}

/** `%{family[0]}`: só o primeiro nome da família, sem caminho nem estilo. */
async function fromFontconfig(): Promise<string[]> {
  const { stdout } = await run('fc-list', ['--format', '%{family[0]}\\n'], {
    timeout: TIMEOUT_MS,
    maxBuffer: MAX_OUTPUT_BYTES,
    windowsHide: true,
  })

  return parseFontconfigFamilies(stdout)
}

/**
 * As duas chaves do registro: a da máquina e a do usuário, onde fica a fonte
 * instalada sem administrador. `reg query` pelo **caminho absoluto**, porque o
 * `CreateProcess` procura o nome simples antes no diretório atual.
 */
async function fromWindowsRegistry(): Promise<string[]> {
  const keys = [
    'HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Fonts',
    'HKCU\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Fonts',
  ]

  const reg = join(process.env['SystemRoot'] ?? 'C:\\Windows', 'System32', 'reg.exe')
  const families = new Set<string>()

  for (const key of keys) {
    try {
      const { stdout } = await run(reg, ['query', key], {
        timeout: TIMEOUT_MS,
        maxBuffer: MAX_OUTPUT_BYTES,
        windowsHide: true,
      })
      for (const family of parseWindowsFontRegistry(stdout)) families.add(family)
    } catch {
      // A chave do usuário não existe em instalação nova.
    }
  }

  return [...families]
}
