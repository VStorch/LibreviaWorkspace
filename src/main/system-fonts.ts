import { execFile } from 'node:child_process'
import { join } from 'node:path'
import { promisify } from 'node:util'
import { parseFontconfigFamilies, parseWindowsFontRegistry } from '@services/document/font-list.js'

/**
 * As fontes que a máquina tem instaladas.
 *
 * O Electron não expõe isso, e nem o Chromium: `queryLocalFonts()` existe no
 * padrão, mas pede permissão do usuário e não vale num renderer sem origem de
 * verdade. Então perguntamos ao sistema, que é trabalho de processo main — o
 * renderer não executa programa, e é justamente essa a garantia do sandbox.
 *
 * A lista é um **conforto**, não um requisito: sem ela a barra continua
 * oferecendo as famílias que viajam no instalador. Por isso toda falha aqui é
 * engolida e devolve lista vazia: nenhum documento deixa de abrir porque o
 * `fc-list` não está instalado.
 */

const run = promisify(execFile)

/**
 * Teto de tempo e de saída.
 *
 * Um `fc-list` num sistema com mil fontes devolve algumas dezenas de milhares de
 * linhas, e é chamado uma vez por sessão. O tempo existe para o caso patológico
 * — cache de fontconfig sendo reconstruído, fonte em disco de rede — em que a
 * barra não pode ficar esperando.
 */
const TIMEOUT_MS = 4000
const MAX_OUTPUT_BYTES = 8 * 1024 * 1024

/**
 * Lida uma vez por sessão.
 *
 * Instalar fonte com o aplicativo aberto é raro; reexecutar um programa a cada
 * vez que a barra de ferramentas monta seria caro e sem ganho nenhum.
 */
let cached: Promise<string[]> | null = null

export function listInstalledFontFamilies(): Promise<string[]> {
  cached ??= collect().then((families) => {
    // Lista vazia **não** fica em cache, e é o único caso em que a leitura se
    // repete. Ela tem duas origens indistinguíveis daqui: máquina sem fontconfig
    // (nada a fazer) e falha transitória — `fc-list` estourando o tempo enquanto o
    // fontconfig reconstrói o cache. Guardar a segunda condenava a sessão inteira
    // a abrir a barra sem fonte nenhuma. Repetir custa um processo por montagem
    // da barra na máquina sem fontconfig; ficar sem lista custa a sessão.
    if (families.length === 0) cached = null
    return families
  })
  return cached
}

/** Só para teste: descarta o que foi lido. Ver `system-fonts.test.ts`. */
export function forgetInstalledFonts(): void {
  cached = null
}

async function collect(): Promise<string[]> {
  try {
    return process.platform === 'win32' ? await fromWindowsRegistry() : await fromFontconfig()
  } catch {
    // Programa ausente, chave de registro inacessível, tempo esgotado: em todos
    // os casos a resposta é a mesma, e ela não é um erro para quem usa.
    return []
  }
}

/**
 * Linux, BSD e macOS com fontconfig instalado.
 *
 * `%{family[0]}` pede só o primeiro nome de cada família — a forma que o CSS
 * acha. Sem o formato, `fc-list` devolve caminho de arquivo e estilo junto, e o
 * seletor mostraria "/usr/share/fonts/… : DejaVu Sans:style=Book".
 */
async function fromFontconfig(): Promise<string[]> {
  const { stdout } = await run('fc-list', ['--format', '%{family[0]}\\n'], {
    timeout: TIMEOUT_MS,
    maxBuffer: MAX_OUTPUT_BYTES,
    windowsHide: true,
  })

  return parseFontconfigFamilies(stdout)
}

/**
 * Windows: as duas chaves de fontes do registro.
 *
 * A da máquina traz as do sistema; a do usuário traz as que ele instalou só para
 * si, que desde o Windows 10 é o caminho padrão de "instalar fonte" sem
 * administrador. Ler só a primeira esconderia justamente as que a pessoa
 * acabou de pôr.
 *
 * `reg query`, e não um módulo de registro: é dependência a menos para auditar,
 * e a saída é estável há décadas. Mas pelo **caminho absoluto**: o `CreateProcess`
 * do Windows procura o nome simples no diretório do executável e no diretório
 * atual antes do `System32`, então um `reg.exe` plantado numa pasta gravável
 * rodaria no lugar do do sistema.
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
      // A chave do usuário não existe em instalação nova. Continuar é o certo:
      // a da máquina sozinha já é uma lista útil.
    }
  }

  return [...families]
}
