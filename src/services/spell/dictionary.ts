/**
 * O que o corretor ortográfico precisa saber sobre o dicionário — em forma pura.
 *
 * O corretor de verdade é o do Chromium, que mora no processo main (ver
 * `src/main/spellcheck.ts`). Aqui ficam só os nomes e a conferência de formato,
 * porque são eles que erram em silêncio: um nome de arquivo diferente do que o
 * Chromium procura faz o corretor **tentar baixar** o dicionário, e numa máquina
 * offline o resultado é nenhuma verificação e nenhum aviso.
 */

/** O único idioma que o aplicativo verifica. Documento em português é o caso. */
export const SPELL_LANGUAGE = 'pt-BR'

/**
 * Revisão do formato binário do Chromium, a parte `-3-0` do nome do arquivo.
 *
 * Não é a versão do dicionário: é a do formato. O Chromium a monta no nome do
 * arquivo que procura, e por isso ela é parte do contrato — não enfeite.
 */
export const DICTIONARY_REVISION = '3-0'

/** Pasta que o Chromium abre dentro do diretório de dados do usuário. */
export const DICTIONARY_FOLDER = 'Dictionaries'

/**
 * Nome do arquivo que o Chromium procura para um idioma.
 *
 * Fonte: o corretor do Chromium monta `<idioma>-<revisão>.bdic` e olha **antes**
 * na pasta local; só baixa quando não acha. É esse "antes" que permite ao
 * aplicativo funcionar sem rede na primeira execução.
 */
export function dictionaryFileName(language: string = SPELL_LANGUAGE): string {
  return `${language}-${DICTIONARY_REVISION}.bdic`
}

/**
 * Assinatura do formato: os quatro bytes `BDic` no começo do arquivo.
 *
 * Um `.bdic` truncado ou trocado por um Hunspell cru (`.dic`, que é texto) é
 * recusado pelo Chromium, que **apaga o arquivo** e parte para o download — o
 * mesmo silêncio que o dicionário embutido existe para evitar. Conferir a
 * assinatura em teste transforma isso em falha de build.
 */
export function hasBdictSignature(bytes: Uint8Array): boolean {
  if (bytes.length < 4) return false
  return bytes[0] === 0x42 && bytes[1] === 0x44 && bytes[2] === 0x69 && bytes[3] === 0x63
}
