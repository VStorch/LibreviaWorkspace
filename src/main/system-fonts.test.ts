import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

/**
 * The system font list on failure, on timeout and per platform. The empty list from an `fc-list`
 * that timed out is not cached. `execFile` is faked in callback style, as `promisify` consumes it.
 */
const { calls, next } = vi.hoisted(() => ({
  calls: [] as Array<{ file: string; args: readonly string[] }>,
  next: [] as Array<{ stdout?: string; error?: Error }>,
}))

vi.mock('node:child_process', () => ({
  execFile: (
    file: string,
    args: readonly string[],
    _options: unknown,
    callback: (error: Error | null, result?: { stdout: string }) => void,
  ) => {
    calls.push({ file, args })
    const reply = next.shift() ?? { stdout: '' }
    if (reply.error !== undefined) callback(reply.error)
    else callback(null, { stdout: reply.stdout ?? '' })
  },
}))

const { forgetInstalledFonts, listInstalledFontFamilies } = await import('./system-fonts.js')

const platform = process.platform

function pretendPlatform(value: string): void {
  Object.defineProperty(process, 'platform', { value, configurable: true })
}

beforeEach(() => {
  calls.length = 0
  next.length = 0
  forgetInstalledFonts()
})

afterEach(() => {
  Object.defineProperty(process, 'platform', { value: platform, configurable: true })
})

describe('fontes instaladas', () => {
  it('no Linux pergunta ao fontconfig e lê uma vez por sessão', async () => {
    pretendPlatform('linux')
    next.push({ stdout: 'DejaVu Sans\nLiberation Serif\n' })

    expect(await listInstalledFontFamilies()).toEqual(['DejaVu Sans', 'Liberation Serif'])
    // The second call runs no program: building the toolbar is frequent.
    expect(await listInstalledFontFamilies()).toEqual(['DejaVu Sans', 'Liberation Serif'])
    expect(calls).toHaveLength(1)
    expect(calls[0]?.file).toBe('fc-list')
  })

  it('a falha não fica em cache', async () => {
    // The point of the test: a timeout is transient, and caching the empty list condemned the whole
    // session to no fonts.
    pretendPlatform('linux')
    next.push({ error: Object.assign(new Error('timeout'), { killed: true }) })

    expect(await listInstalledFontFamilies()).toEqual([])

    next.push({ stdout: 'Carlito\n' })
    expect(await listInstalledFontFamilies()).toEqual(['Carlito'])
    expect(calls).toHaveLength(2)
  })

  it('no Windows chama o reg.exe do System32, pelo caminho absoluto', async () => {
    // `reg` by its bare name is looked up in the executable's directory and the current directory
    // before System32: a `reg.exe` planted in a writable folder would run instead of the system's.
    pretendPlatform('win32')
    process.env['SystemRoot'] = 'C:\\Windows'
    next.push({ stdout: '    Arial (TrueType)    REG_SZ    arial.ttf\n' })
    next.push({ error: new Error('a chave do usuário não existe') })

    expect(await listInstalledFontFamilies()).toEqual(['Arial'])
    expect(calls).toHaveLength(2)
    // The separator is the test machine's: on real Windows `join` writes backslashes. What matters
    // is the full path.
    const reg = join('C:\\Windows', 'System32', 'reg.exe')
    for (const call of calls) expect(call.file).toBe(reg)
    // Both keys: the machine's and the user's, where Windows 10 puts fonts installed without admin
    // rights.
    expect(calls[0]?.args[1]).toContain('HKLM')
    expect(calls[1]?.args[1]).toContain('HKCU')
  })
})
