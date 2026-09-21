import { closeSync, copyFileSync, mkdirSync, openSync, readFileSync, readSync } from 'node:fs'
import { readFile } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { app, protocol, type Session } from 'electron'
import {
  DICTIONARY_FOLDER,
  SPELL_LANGUAGE,
  dictionaryFileName,
  hasBdictSignature,
} from '@services/spell/dictionary.js'
import { DictionaryScope } from '@shared/types.js'

/**
 * Verificação ortográfica em português — **sem rede, nem na primeira execução**.
 *
 * O corretor é o do Chromium, que vem embutido no Electron. O problema não é o
 * corretor: é o dicionário. Deixado por conta própria, o Chromium o baixa de um
 * CDN quando o idioma é escolhido, e numa máquina sem rede o resultado é nenhuma
 * verificação e nenhum aviso — a falha silenciosa que este projeto trata como
 * defeito grave.
 *
 * ## Por que copiar, e não apontar
 *
 * Duas APIs parecem servir e não servem:
 *
 *  - `setSpellCheckerDictionaryDownloadURL` continua sendo um **download**: ela
 *    troca o endereço, não a natureza da operação. Um `file:` ali não é buscado
 *    pela pilha de rede do Chromium, e um servidor local seria justamente o que o
 *    aplicativo se recusa a ter;
 *  - `webFrame.setSpellCheckProvider` substituiria o corretor inteiro, e aí o
 *    dicionário, as sugestões e a marcação vermelha passariam a ser nossos —
 *    trabalho de outra ordem para um resultado pior.
 *
 * O que o Chromium faz **antes** de baixar é procurar o arquivo em disco, na
 * pasta `Dictionaries` do diretório de dados do usuário, com um nome que ele
 * mesmo monta (`pt-BR-3-0.bdic`). Então a solução é simplesmente pôr o arquivo
 * ali na primeira execução: ele acha, carrega e nunca pergunta pela rede.
 *
 * O `setSpellCheckerDictionaryDownloadURL` ainda é chamado, apontando para um
 * esquema nosso — não como caminho principal, mas como cinto de segurança: se um
 * dia o nome que o Chromium procura mudar, o pedido cai no nosso protocolo (que
 * serve o mesmo arquivo do pacote) e, mesmo falhando, não vaza para a internet.
 */

/** Esquema que responde ao downloader do corretor. Ver o comentário acima. */
export const DICTIONARY_SCHEME = 'librevia-dict'

/**
 * Palavras que valem só nesta sessão — o "ignorar" do menu de contexto.
 *
 * O Chromium não tem lista de ignorados: ou a palavra está no dicionário do
 * usuário, ou é marcada. "Ignorar" é então uma entrada no dicionário que é
 * desfeita ao sair, e é por isso que ela precisa ser lembrada aqui.
 */
const sessionWords = new Set<string>()

/**
 * Onde o dicionário embutido mora — mesma regra das fontes e do sidecar.
 *
 * `app.getAppPath()` muda conforme o Electron é chamado, e essa variação já
 * custou um diagnóstico errado uma vez (ver `src/main/fonts.ts`).
 */
function bundledDictionaryPath(): string {
  const root = app.isPackaged
    ? process.resourcesPath
    : join(dirname(fileURLToPath(import.meta.url)), '..', '..')
  return join(root, 'resources', 'dictionaries', dictionaryFileName())
}

/** A pasta em que o Chromium procura o dicionário antes de pensar em baixar. */
function installedDictionaryPath(): string {
  return join(app.getPath('userData'), DICTIONARY_FOLDER, dictionaryFileName())
}

/**
 * Põe o dicionário embutido no lugar onde o Chromium o procura.
 *
 * Idempotente: se o arquivo já está lá — de uma execução anterior ou baixado por
 * uma versão antiga do aplicativo — nada é feito. Devolve `false` quando não deu
 * para instalar, e nesse caso a ortografia simplesmente não vai marcar nada; é o
 * único desfecho em que perder o recurso é melhor que impedir o aplicativo de
 * abrir.
 *
 * Precisa rodar antes de a sessão padrão existir. Ver "Quando, e por que
 * síncrono", acima.
 */
export function installBundledDictionary(): boolean {
  // No macOS o corretor é o do sistema e não existe `.bdic` nenhum: copiar o
  // arquivo ali seria lixo no perfil do usuário.
  if (process.platform === 'darwin') return true

  const target = installedDictionaryPath()

  try {
    // Já está lá — de uma execução anterior, ou baixado por uma versão do
    // aplicativo que ainda não levava o dicionário embutido. Mas estar lá não
    // basta: **a assinatura** decide. Um `.bdic` truncado ou estragado é apagado
    // pelo Chromium, que então tentaria baixar, e numa máquina sem rede a sessão
    // fica sem ortografia sem avisar ninguém — a falha silenciosa de sempre.
    // Conferindo aqui, o arquivo é reposto nesta execução, e não na seguinte.
    if (hasInstalledSignature(target)) return true
    console.error(`[spellcheck] dicionário do perfil está corrompido e será reposto: ${target}`)
  } catch {
    // Não existe ainda: é o caminho normal da primeira execução.
  }

  try {
    const source = bundledDictionaryPath()

    // Um `.bdic` inválido é apagado pelo Chromium, que então tenta baixar —
    // exatamente o que não pode acontecer. Melhor descobrir aqui, no log, do que
    // ficar sem corretor sem saber por quê.
    if (!hasBdictSignature(readFileSync(source))) {
      console.error(`[spellcheck] dicionário embutido não está no formato BDic: ${source}`)
      return false
    }

    mkdirSync(dirname(target), { recursive: true })
    copyFileSync(source, target)
    return true
  } catch (cause) {
    console.error('[spellcheck] não foi possível instalar o dicionário embutido:', cause)
    return false
  }
}

/**
 * A assinatura do arquivo que está no perfil, lendo só os quatro primeiros bytes.
 *
 * Quatro, e não o arquivo inteiro: isto roda a cada abertura, de forma síncrona e
 * antes de a sessão existir (ver "Quando, e por que síncrono"), e o dicionário de
 * português tem alguns megabytes. O que se quer saber está no cabeçalho.
 */
function hasInstalledSignature(target: string): boolean {
  const handle = openSync(target, 'r')
  try {
    const head = new Uint8Array(4)
    readSync(handle, head, 0, 4, 0)
    return hasBdictSignature(head)
  } finally {
    closeSync(handle)
  }
}

/**
 * Liga o esquema que responde ao downloader do corretor.
 *
 * Chegar aqui significa que o arquivo local não foi encontrado — daí o aviso no
 * log. Servir o mesmo arquivo do pacote é a tentativa de salvar a situação sem
 * sair da máquina.
 */
export function serveDictionary(): void {
  protocol.handle(DICTIONARY_SCHEME, async (request) => {
    console.warn(`[spellcheck] o corretor pediu o dicionário pela rede: ${request.url}`)

    try {
      return new Response(await readFile(bundledDictionaryPath()), {
        headers: { 'Content-Type': 'application/octet-stream' },
      })
    } catch {
      return new Response('', { status: 404 })
    }
  })
}

/**
 * Aplica o estado da preferência à sessão.
 *
 * A ordem importa: o endereço de download é definido **antes** do idioma, porque
 * é escolher o idioma que dispara a procura pelo dicionário.
 *
 * Numa máquina cujo idioma padrão não seja o português, o Chromium pode tentar
 * uma vez o dicionário do idioma dele antes de obedecer a esta chamada — o padrão
 * dele vem do sistema, não de nós. Sem rede a tentativa falha e nada acontece; o
 * idioma que passa a valer é o pt-BR daqui, cujo dicionário já está em disco.
 */
export function applySpellChecker(session: Session, enabled: boolean): void {
  try {
    session.setSpellCheckerDictionaryDownloadURL(`${DICTIONARY_SCHEME}://dictionaries/`)

    if (enabled) session.setSpellCheckerLanguages([SPELL_LANGUAGE])
    session.setSpellCheckerEnabled(enabled)
  } catch (cause) {
    // Um Electron compilado sem corretor (ou o macOS sem o idioma instalado)
    // lança daqui. Perder a ortografia é aceitável; derrubar o aplicativo na
    // inicialização, não.
    console.error('[spellcheck] o corretor não pôde ser configurado:', cause)
  }
}

/**
 * Guarda a palavra no dicionário do usuário.
 *
 * `session` é o "ignorar": entra no dicionário agora e sai na saída do
 * aplicativo. Sem isso, "ignorar" e "adicionar ao dicionário" seriam o mesmo
 * botão com dois nomes.
 */
export function rememberWord(session: Session, word: string, scope: DictionaryScope): boolean {
  const added = session.addWordToSpellCheckerDictionary(word)
  if (added && scope === DictionaryScope.Session) sessionWords.add(word)
  return added
}

/** Desfaz os "ignorar" desta sessão. Chamado ao encerrar. */
export function forgetSessionWords(session: Session): void {
  for (const word of sessionWords) {
    try {
      session.removeWordFromSpellCheckerDictionary(word)
    } catch {
      // O dicionário do usuário é conveniência: falhar em limpá-lo não pode
      // atrasar o encerramento.
    }
  }
  sessionWords.clear()
}
