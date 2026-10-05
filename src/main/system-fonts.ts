import { execFile } from 'node:child_process'
import { join } from 'node:path'
import { promisify } from 'node:util'
import { parseFontconfigFamilies, parseWindowsFontRegistry } from '@services/document/font-list.js'

/**
 * `queryLocalFonts()` asks for permission and does not work in a renderer without an origin, so
 * main asks. A comfort, not a requirement: every failure returns an empty list.
 */

const run = promisify(execFile)

/** For the pathological case, such as the fontconfig cache being rebuilt. */
const TIMEOUT_MS = 4000
const MAX_OUTPUT_BYTES = 8 * 1024 * 1024

/** Once per session: installing a font with the app open is rare. */
let cached: Promise<string[]> | null = null

export function listInstalledFontFamilies(): Promise<string[]> {
  cached ??= collect().then((families) => {
    // An empty list is not cached: it may be a transient failure, and keeping it would leave the
    // toolbar without fonts for the whole session.
    if (families.length === 0) cached = null
    return families
  })
  return cached
}

/** For the tests. */
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

/** `%{family[0]}`: only the first family name, without path or style. */
async function fromFontconfig(): Promise<string[]> {
  const { stdout } = await run('fc-list', ['--format', '%{family[0]}\\n'], {
    timeout: TIMEOUT_MS,
    maxBuffer: MAX_OUTPUT_BYTES,
    windowsHide: true,
  })

  return parseFontconfigFamilies(stdout)
}

/**
 * Both registry keys: the machine's and the user's, where fonts installed without admin rights go.
 * `reg query` by **absolute path**, because `CreateProcess` looks up the bare name in the current
 * directory first.
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
      // The user key does not exist on a fresh install.
    }
  }

  return [...families]
}
