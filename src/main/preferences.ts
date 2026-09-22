import { app, nativeTheme, session } from 'electron'
import Store from 'electron-store'
import { IpcChannel } from '@shared/ipc-channels.js'
import { languageFromLocale } from '@shared/i18n/index.js'
import { editorPreferencesSchema } from '@shared/schemas.js'
import {
  DEFAULT_EDITOR_PREFERENCES,
  Theme,
  type EditorPreferences,
  type EditorPreferencesPatch,
  type ResolvedTheme,
} from '@shared/types.js'
import { applySpellChecker } from './spellcheck.js'
import { broadcastPush } from './window.js'

/**
 * As preferências de edição, e quem manda nelas.
 *
 * O main é o dono por uma razão concreta: a ortografia é configuração de `session`,
 * e só ele fala com o corretor do Chromium. Se o renderer guardasse a sua metade,
 * o item do menu diria "ligado" enquanto o editor não marcava nada — ou o
 * contrário.
 *
 * Todo caminho de mudança passa por `updatePreferences`: o menu nativo, a barra
 * de ferramentas e o que um dia mais aparecer. É ele que grava, aplica na sessão
 * e avisa quem precisa saber.
 */

interface PreferencesSchema {
  preferences: EditorPreferences
}

const store = new Store<PreferencesSchema>({
  name: 'preferences',
  defaults: { preferences: DEFAULT_EDITOR_PREFERENCES },
  // Mesmo critério da lista de recentes: um JSON corrompido não pode impedir o
  // aplicativo de abrir. Preferência é conveniência, não dado do usuário.
  clearInvalidConfig: true,
})

/**
 * O estado vive em memória e o arquivo é só a cópia durável.
 *
 * Ler do disco a cada consulta seria trabalho por nada — o menu é reconstruído
 * várias vezes por sessão — e, pior, faria o valor depender de a gravação ter
 * terminado.
 */
let current: EditorPreferences = load()

const listeners = new Set<(preferences: EditorPreferences) => void>()

function load(): EditorPreferences {
  const stored = store.get('preferences')

  // Pelo schema, e não pelo que estiver no arquivo: os `default` dele são o que
  // permite abrir um perfil gravado por uma versão que não tinha estas chaves.
  const parsed = editorPreferencesSchema.safeParse(stored)
  const preferences = parsed.success ? parsed.data : DEFAULT_EDITOR_PREFERENCES

  // O idioma da primeira execução sai do sistema operacional, e não do
  // `default` do schema.
  //
  // A diferença aparece exatamente uma vez na vida de uma instalação, e é a
  // diferença entre um programa que abre na língua da pessoa e um que abre em
  // português e a obriga a procurar onde se troca. Por isso a pergunta não é
  // "qual é o valor?" — que o `default` já responde — mas "a chave foi
  // gravada?". Um `default` apaga essa distinção, então ela é lida do objeto
  // cru, antes do parse.
  //
  // Depois da primeira gravação a chave existe, e a escolha da pessoa vale
  // mesmo que ela troque a língua do sistema depois.
  if (declares(stored, 'language')) return preferences
  return { ...preferences, language: languageFromLocale(app.getLocale()) }
}

/** A chave estava no arquivo, em vez de ter vindo do `default` do schema. */
function declares(stored: unknown, key: string): boolean {
  return typeof stored === 'object' && stored !== null && key in stored
}

/**
 * O tema que a tela deve desenhar, com `system` já resolvido.
 *
 * `nativeTheme.shouldUseDarkColors` é a resposta do Chromium à configuração do
 * sistema, e é ela que muda sozinha quando a pessoa troca de claro para escuro
 * sem fechar o aplicativo. Uma escolha explícita ignora o sistema — é o que
 * "explícita" quer dizer.
 */
export function resolvedTheme(): ResolvedTheme {
  if (current.theme === Theme.Light) return 'light'
  if (current.theme === Theme.Dark) return 'dark'
  return nativeTheme.shouldUseDarkColors ? 'dark' : 'light'
}

export function editorPreferences(): EditorPreferences {
  return current
}

/**
 * Liga o que está guardado na sessão do Chromium.
 *
 * Chamado uma vez na inicialização, depois de o dicionário embutido estar no
 * lugar — ver `installBundledDictionary`.
 */
export function applyStoredPreferences(): void {
  applySpellChecker(session.defaultSession, current.spellcheck)
  // Antes de a janela abrir: assim a primeira pintura já sai na cor certa, em
  // vez de mostrar o tema claro por um quadro e trocar depois.
  nativeTheme.themeSource = current.theme
}

/**
 * Muda uma ou mais preferências e devolve o estado resultante.
 *
 * Remendo, e não conjunto inteiro: quem clica em "marcas de formatação" não tem
 * opinião sobre ortografia.
 */
export function updatePreferences(patch: EditorPreferencesPatch): EditorPreferences {
  // Chave por chave, e não por espalhamento: o remendo pode trazer a chave
  // presente com `undefined`, e espalhá-la apagaria a preferência em vez de
  // deixá-la como estava.
  const next: EditorPreferences = {
    spellcheck: patch.spellcheck ?? current.spellcheck,
    invisibleCharacters: patch.invisibleCharacters ?? current.invisibleCharacters,
    typography: patch.typography ?? current.typography,
    language: patch.language ?? current.language,
    theme: patch.theme ?? current.theme,
    readingMode: patch.readingMode ?? current.readingMode,
  }

  const spellcheckChanged = next.spellcheck !== current.spellcheck
  const themeChanged = next.theme !== current.theme

  // Chave a chave sobre o próprio tipo, e não uma lista escrita à mão: a lista
  // anterior tinha de crescer a cada preferência nova, e a que alguém
  // esquecesse de acrescentar passaria a ser gravada sem avisar ninguém —
  // falha silenciosa, do tipo que só aparece semanas depois.
  const keys = Object.keys(next) as (keyof EditorPreferences)[]
  const unchanged = keys.every((key) => next[key] === current[key])

  // Sem isto, reabrir o menu com o mesmo valor gravaria o arquivo e mandaria um
  // aviso — e o aviso redesenharia o editor por nada.
  if (unchanged) return current

  current = next
  store.set('preferences', next)

  if (spellcheckChanged) applySpellChecker(session.defaultSession, next.spellcheck)

  // O renderer resolve `system` sozinho, por `matchMedia`, e é esta linha que
  // faz isso funcionar: o `themeSource` do Chromium é o que a consulta de mídia
  // enxerga. Assim o tema não precisa de canal de IPC próprio nem de um segundo
  // valor "resolvido" viajando junto das preferências — e, em `system`, mudar o
  // tema do sistema operacional chega à tela sem o main fazer nada.
  if (themeChanged) nativeTheme.themeSource = next.theme

  // O renderer é avisado sempre, inclusive quando foi ele quem pediu: é assim
  // que a barra de ferramentas e o menu nativo mostram a mesma coisa.
  broadcastPush(IpcChannel.PreferencesChanged, next)
  for (const listener of listeners) listener(next)

  return next
}

/**
 * Avisa quem precisa refazer algo no main — hoje, o menu, para o item de
 * marcação acompanhar o estado.
 *
 * Um emissor, e não uma chamada direta a `refreshMenu`: o menu já importa estas
 * preferências para desenhar as marcas de seleção, e importar um ao outro faria
 * um ciclo.
 */
export function onPreferencesChanged(listener: (preferences: EditorPreferences) => void): void {
  listeners.add(listener)
}
