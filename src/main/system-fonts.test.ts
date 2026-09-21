import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

/**
 * A lista de fontes do sistema, nos três caminhos que ninguém cobria: a falha, o
 * tempo esgotado e a escolha de plataforma.
 *
 * O caso que motivou o teste é o silencioso: `fc-list` estourando o tempo durante
 * uma reconstrução do cache do fontconfig devolvia lista vazia — e a lista vazia
 * ficava em cache pela sessão inteira, então a barra abria sem fonte nenhuma até o
 * aplicativo ser reiniciado.
 *
 * `execFile` é falsificado no estilo de callback, que é como o `promisify` do
 * módulo o consome.
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
    // A segunda chamada não executa programa nenhum: montar a barra é frequente.
    expect(await listInstalledFontFamilies()).toEqual(['DejaVu Sans', 'Liberation Serif'])
    expect(calls).toHaveLength(1)
    expect(calls[0]?.file).toBe('fc-list')
  })

  it('a falha não fica em cache', async () => {
    // O ponto do teste: tempo esgotado é transitório, e guardar a lista vazia
    // condenava a sessão inteira a ficar sem fontes.
    pretendPlatform('linux')
    next.push({ error: Object.assign(new Error('timeout'), { killed: true }) })

    expect(await listInstalledFontFamilies()).toEqual([])

    next.push({ stdout: 'Carlito\n' })
    expect(await listInstalledFontFamilies()).toEqual(['Carlito'])
    expect(calls).toHaveLength(2)
  })

  it('no Windows chama o reg.exe do System32, pelo caminho absoluto', async () => {
    // `reg` pelo nome simples é procurado no diretório do executável e no diretório
    // atual antes do System32: um `reg.exe` plantado numa pasta gravável seria
    // executado no lugar do do sistema.
    pretendPlatform('win32')
    process.env['SystemRoot'] = 'C:\\Windows'
    next.push({ stdout: '    Arial (TrueType)    REG_SZ    arial.ttf\n' })
    next.push({ error: new Error('a chave do usuário não existe') })

    expect(await listInstalledFontFamilies()).toEqual(['Arial'])
    expect(calls).toHaveLength(2)
    // O separador é o da máquina que roda o teste — no Windows de verdade o
    // `join` escreve com barra invertida. O que importa é o caminho completo.
    const reg = join('C:\\Windows', 'System32', 'reg.exe')
    for (const call of calls) expect(call.file).toBe(reg)
    // As duas chaves: a da máquina e a do usuário, que é onde o Windows 10 põe
    // fonte instalada sem administrador.
    expect(calls[0]?.args[1]).toContain('HKLM')
    expect(calls[1]?.args[1]).toContain('HKCU')
  })
})
