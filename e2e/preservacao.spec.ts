import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { expect, test } from '@playwright/test'
import { launch, menu, stubDialogs, type Session } from './app.js'
import { docxWithTextBox } from './fixtures.js'

/**
 * Abrir e salvar sem editar não pode mexer no arquivo.
 *
 * A gravação é cirúrgica: cada bloco carrega um `oid` do leitor, e o gravador
 * devolve o XML original de todo bloco cujo conteúdo não mudou. É isso que faz
 * comentário, revisão, caixa de texto e forma sobreviverem num documento que o
 * editor não sabe reproduzir.
 *
 * O caminho inteiro passa pelo ProseMirror, e é lá que ele quebrava: o `oid`
 * não estava declarado no schema e era descartado na travessia, e a comparação
 * que decide o que preservar reprovava blocos por diferenças de forma — atributo
 * nulo, ordem de chave, ordem de marca, texto partido por run. As duas coisas
 * juntas faziam **abrir e salvar** regenerar o documento inteiro em silêncio.
 *
 * Nenhum teste pegava: os do sidecar vão do leitor ao gravador sem passar pelo
 * editor, que é justamente o trecho onde a identidade se perdia. Este vai pelo
 * aplicativo montado, que é o único lugar onde a travessia existe de verdade.
 */
test.describe('gravação cirúrgica', () => {
  let session: Session
  let folder: string

  test.beforeEach(async () => {
    folder = await mkdtemp(join(tmpdir(), 'librevia-preservacao-'))
    session = await launch()
  })

  test.afterEach(async () => {
    await session.close()
    await rm(folder, { recursive: true, force: true })
  })

  test('abrir e salvar sem editar preserva a caixa de texto', async () => {
    const origem = join(folder, 'entrada.docx')
    const destino = join(folder, 'saida.docx')
    await writeFile(origem, await docxWithTextBox())

    await stubDialogs(session.app, { open: origem, save: destino, messageBox: 1 })
    await menu(session, 'open')
    await expect(session.window.locator('.ProseMirror')).toContainText('Título na caixa')

    await menu(session, 'save-as')
    await expect(session.window.locator('.statusbar__state')).toHaveText('Salvo')

    // A caixa é a sentinela: o editor mostra o texto de dentro, mas não sabe
    // redesenhar a forma. Se ela voltou, o XML original do bloco foi preservado
    // — que é a única coisa que este teste precisa saber.
    const corpo = await corpoDoDocumento(destino)
    expect(corpo).toContain('txbxContent')
    expect(corpo).toContain('Título na caixa')
  })
})

/** Conteúdo de `word/document.xml` dentro do `.docx`, sem descompactar em disco. */
async function corpoDoDocumento(caminho: string): Promise<string> {
  const { promisify } = await import('node:util')
  const { inflateRaw } = await import('node:zlib')
  const inflate = promisify(inflateRaw)
  const zip = await readFile(caminho)

  // Varredura dos cabeçalhos locais do zip: o suficiente para achar uma parte
  // pelo nome, e sem trazer dependência nova para os testes.
  for (let i = 0; i + 30 <= zip.length; i++) {
    if (zip.readUInt32LE(i) !== 0x04034b50) continue

    const metodo = zip.readUInt16LE(i + 8)
    const comprimido = zip.readUInt32LE(i + 18)
    const original = zip.readUInt32LE(i + 22)
    const tamanhoNome = zip.readUInt16LE(i + 26)
    const extra = zip.readUInt16LE(i + 28)
    const nome = zip.subarray(i + 30, i + 30 + tamanhoNome).toString('utf8')
    if (nome !== 'word/document.xml') continue

    const fim = i + 30 + tamanhoNome + extra + (comprimido > 0 ? comprimido : original)
    const dados = zip.subarray(i + 30 + tamanhoNome + extra, fim)
    return (metodo === 0 ? dados : await inflate(dados)).toString('utf8')
  }

  throw new Error(`word/document.xml não encontrado em ${caminho}`)
}
