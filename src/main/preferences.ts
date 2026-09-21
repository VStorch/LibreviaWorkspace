import { session } from 'electron'
import Store from 'electron-store'
import { IpcChannel } from '@shared/ipc-channels.js'
import { editorPreferencesSchema } from '@shared/schemas.js'
import {
  DEFAULT_EDITOR_PREFERENCES,
  type EditorPreferences,
  type EditorPreferencesPatch,
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
  // Pelo schema, e não pelo que estiver no arquivo: os `default` dele são o que
  // permite abrir um perfil gravado por uma versão que não tinha estas chaves.
  const parsed = editorPreferencesSchema.safeParse(store.get('preferences'))
  return parsed.success ? parsed.data : DEFAULT_EDITOR_PREFERENCES
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
  }

  const spellcheckChanged = next.spellcheck !== current.spellcheck
  const unchanged =
    !spellcheckChanged &&
    next.invisibleCharacters === current.invisibleCharacters &&
    next.typography === current.typography

  // Sem isto, reabrir o menu com o mesmo valor gravaria o arquivo e mandaria um
  // aviso — e o aviso redesenharia o editor por nada.
  if (unchanged) return current

  current = next
  store.set('preferences', next)

  if (spellcheckChanged) applySpellChecker(session.defaultSession, next.spellcheck)

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
