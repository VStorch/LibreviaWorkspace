import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { afterAll, describe, expect, it, vi } from 'vitest'
import { DICTIONARY_FOLDER, dictionaryFileName, hasBdictSignature } from '@services/spell/dictionary.js'

const raiz = join(dirname(fileURLToPath(import.meta.url)), '..', '..')

/**
 * Um perfil de usuário de mentira, para o caminho de instalação do dicionário.
 *
 * `app.getPath('userData')` é a única coisa de que `installBundledDictionary`
 * precisa do Electron — e `app.isPackaged` falso, que é o que faz o arquivo
 * embutido ser procurado na árvore do projeto.
 */
const perfil = mkdtempSync(join(tmpdir(), 'librevia-spell-'))

vi.mock('electron', () => ({
  app: { isPackaged: false, getPath: () => perfil },
  protocol: { handle: () => {} },
}))

const { installBundledDictionary } = await import('./spellcheck.js')

afterAll(() => {
  rmSync(perfil, { recursive: true, force: true })
})

describe('dicionário embutido', () => {
  it('existe em resources/dictionaries e está no formato do Chromium', () => {
    // Mesmo motivo do teste das fontes (`src/main/fonts.test.ts`): sem ele a
    // falha é silenciosa e só aparece na máquina de quem instalou — que não tem
    // dicionário nenhum no perfil e, sem rede, fica sem corretor.
    const arquivo = join(raiz, 'resources', DICTIONARY_FOLDER.toLowerCase(), dictionaryFileName())
    const conteudo = readFileSync(arquivo)

    expect(hasBdictSignature(conteudo)).toBe(true)
    // Um dicionário de português inteiro tem alguns megabytes. O piso pega o caso
    // de o arquivo ter virado um marcador de Git LFS ou um download interrompido.
    expect(conteudo.byteLength).toBeGreaterThan(1_000_000)
  })
})

describe('instalação do dicionário no perfil', () => {
  const instalado = join(perfil, DICTIONARY_FOLDER, dictionaryFileName())

  it('copia o dicionário embutido na primeira execução', () => {
    expect(installBundledDictionary()).toBe(true)
    expect(hasBdictSignature(readFileSync(instalado))).toBe(true)
  })

  it('repara o arquivo estragado na mesma execução', () => {
    // O modo de falha que o QA reproduziu: um `.bdic` corrompido no perfil — de um
    // download interrompido por uma versão antiga, ou de disco cheio — é apagado
    // pelo Chromium, que então tentaria baixar. Numa máquina sem rede a sessão
    // fica sem ortografia **sem avisar**, que é o defeito que este módulo existe
    // para evitar. Conferir só o tamanho deixava o arquivo de pé até a próxima
    // execução; a assinatura o repõe agora.
    writeFileSync(instalado, 'lixo')

    expect(installBundledDictionary()).toBe(true)
    expect(hasBdictSignature(readFileSync(instalado))).toBe(true)
  })
})
