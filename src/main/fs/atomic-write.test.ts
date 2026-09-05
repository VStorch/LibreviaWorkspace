import { chmod, mkdtemp, readFile, readdir, rm, stat, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { AppError } from '@shared/errors.js'
import { writeFileAtomic } from './atomic-write.js'

/**
 * Três testes abaixo afirmam coisas que só existem em POSIX.
 *
 * No Windows o `chmod` do Node não escreve permissão nenhuma: ele liga e
 * desliga o atributo de somente leitura, e `stat` devolve 0o666 para qualquer
 * arquivo gravável — daí `expected 438 to be 416`. Em diretório ele não faz
 * efeito algum: a pasta continua gravável, a gravação dá certo e o teste que
 * espera falha não vê falha nenhuma. Não é o `writeFileAtomic` que muda de
 * comportamento; é a permissão que não existe do outro lado, e forçar uma
 * asserção equivalente exigiria mexer em ACL, que não é o que este módulo faz.
 *
 * O que eles protegem — não estreitar o acesso de um arquivo compartilhado, e
 * não destruir o original quando a gravação falha — continua coberto no Linux,
 * que é onde a permissão significa alguma coisa.
 */
const emPosix = process.platform !== 'win32'

let directory: string

beforeEach(async () => {
  directory = await mkdtemp(join(tmpdir(), 'librevia-test-'))
})

afterEach(async () => {
  await chmod(directory, 0o700).catch(() => undefined)
  await rm(directory, { recursive: true, force: true })
})

describe('writeFileAtomic', () => {
  it('cria um arquivo novo', async () => {
    const target = join(directory, 'novo.txt')
    await writeFileAtomic(target, 'conteúdo inicial')

    expect(await readFile(target, 'utf8')).toBe('conteúdo inicial')
  })

  it('não deixa arquivo temporário para trás', async () => {
    await writeFileAtomic(join(directory, 'a.txt'), 'x')

    // Um .tmp esquecido na pasta do usuário é lixo visível — e, numa pasta de
    // rede compartilhada, lixo que todo mundo vê.
    const leftovers = (await readdir(directory)).filter((name) => name.endsWith('.tmp'))
    expect(leftovers).toEqual([])
  })

  it('guarda o conteúdo anterior em .bak ao sobrescrever', async () => {
    const target = join(directory, 'ata.txt')
    await writeFileAtomic(target, 'versão 1')
    await writeFileAtomic(target, 'versão 2')

    expect(await readFile(target, 'utf8')).toBe('versão 2')
    // O .bak precisa ter a versão *anterior*: é essa a proteção contra uma
    // gravação equivocada por cima de um documento bom.
    expect(await readFile(`${target}.bak`, 'utf8')).toBe('versão 1')
  })

  it('não cria .bak quando o arquivo ainda não existia', async () => {
    const target = join(directory, 'primeiro.txt')
    await writeFileAtomic(target, 'conteúdo')

    await expect(stat(`${target}.bak`)).rejects.toThrow()
  })

  it.runIf(emPosix)('preserva as permissões do arquivo existente', async () => {
    const target = join(directory, 'compartilhado.txt')
    await writeFile(target, 'original')
    await chmod(target, 0o640)

    await writeFileAtomic(target, 'atualizado')

    // Salvar não pode estreitar o acesso de um arquivo compartilhado por uma
    // equipe: quem podia ler antes precisa continuar podendo.
    expect((await stat(target)).mode & 0o777).toBe(0o640)
  })

  it.runIf(emPosix)('mantém o arquivo original intacto quando a gravação falha', async () => {
    const target = join(directory, 'protegido.txt')
    await writeFile(target, 'conteúdo valioso')
    await chmod(directory, 0o500) // leitura e travessia, sem escrita

    await expect(writeFileAtomic(target, 'tentativa')).rejects.toBeInstanceOf(AppError)

    await chmod(directory, 0o700)
    expect(await readFile(target, 'utf8')).toBe('conteúdo valioso')
  })

  it.runIf(emPosix)('reporta falha de gravação com mensagem compreensível', async () => {
    await chmod(directory, 0o500)

    await expect(writeFileAtomic(join(directory, 'x.txt'), 'a')).rejects.toMatchObject({
      message: expect.stringMatching(/permissão|salvar/i),
    })
  })

  it('grava num sistema de arquivos diferente do temporário do sistema', async () => {
    // Este é o teste da pasta de rede. Uma implementação que preparasse o
    // arquivo temporário em os.tmpdir() e depois renomeasse para o destino
    // falharia aqui com EXDEV — que é exatamente o que acontece ao salvar num
    // compartilhamento montado. Só passa porque o temporário nasce na pasta
    // de destino.
    const crossDevice = await mkdtemp(join('/dev/shm', 'librevia-xdev-')).catch(() => null)
    if (crossDevice === null) return // ambiente sem tmpfs separado

    try {
      const target = join(crossDevice, 'em-outro-volume.txt')
      await writeFileAtomic(target, 'primeira versão')
      await writeFileAtomic(target, 'segunda versão')

      expect(await readFile(target, 'utf8')).toBe('segunda versão')
      expect(await readFile(`${target}.bak`, 'utf8')).toBe('primeira versão')
    } finally {
      await rm(crossDevice, { recursive: true, force: true })
    }
  })

  it('grava conteúdo com acentuação e quebras de linha sem alterar bytes', async () => {
    const target = join(directory, 'acentos.txt')
    const content = 'Ação\nCoração — “aspas”\r\nfim\t.'
    await writeFileAtomic(target, content)

    expect(await readFile(target, 'utf8')).toBe(content)
  })
})
