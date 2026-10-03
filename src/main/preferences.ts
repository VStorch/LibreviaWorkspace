import { userInfo } from 'node:os'
import * as electron from 'electron'
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
 * O main é o dono porque a ortografia é configuração de `session`. Todo caminho
 * de mudança passa por `updatePreferences`, que grava, aplica e avisa.
 */

interface PreferencesSchema {
  preferences: EditorPreferences
}

let storeInstance: Store<PreferencesSchema> | null = null

function getStore(): Store<PreferencesSchema> {
  if (storeInstance === null) {
    storeInstance = new Store<PreferencesSchema>({
      name: 'preferences',
      defaults: { preferences: DEFAULT_EDITOR_PREFERENCES },
      // Um JSON corrompido não impede o aplicativo de abrir: preferência é conveniência.
      clearInvalidConfig: true,
      ...(process.versions.electron ? {} : { cwd: process.cwd() }),
    })
  }
  return storeInstance
}

/** Em memória; o arquivo é a cópia durável. */
let current: EditorPreferences | null = null

const listeners = new Set<(preferences: EditorPreferences) => void>()

function load(): EditorPreferences {
  return withAuthor(loadStored())
}

/** Sem autor, o usuário do sistema, como o Word faz na primeira vez. */
function withAuthor(preferences: EditorPreferences): EditorPreferences {
  if (preferences.authorName.trim() !== '') return preferences
  let username = ''
  try {
    username = userInfo().username
  } catch {
    // Sem conta no sistema (contêiner), o comentário sai sem autor.
  }
  return { ...preferences, authorName: username }
}

function loadStored(): EditorPreferences {
  if (!process.versions.electron) return DEFAULT_EDITOR_PREFERENCES

  try {
    const store = getStore()
    const stored = store.get('preferences')

    // Pelo schema: os `default` abrem o perfil gravado sem estas chaves.
    const parsed = editorPreferencesSchema.safeParse(stored)
    const preferences = parsed.success ? parsed.data : DEFAULT_EDITOR_PREFERENCES

    // Na primeira execução o idioma vem do sistema operacional. A pergunta é "a
    // chave foi gravada?", e o `default` apagaria essa distinção, por isso ela é
    // lida do objeto cru.
    if (declares(stored, 'language')) return preferences
    const locale = typeof electron.app?.getLocale === 'function' ? electron.app.getLocale() : 'pt-BR'
    return { ...preferences, language: languageFromLocale(locale) }
  } catch {
    return DEFAULT_EDITOR_PREFERENCES
  }
}

function declares(stored: unknown, key: string): boolean {
  return typeof stored === 'object' && stored !== null && key in stored
}

/** `nativeTheme.shouldUseDarkColors` muda sozinho quando a pessoa troca o tema do sistema. */
export function resolvedTheme(): ResolvedTheme {
  const preferences = editorPreferences()
  if (preferences.theme === Theme.Light) return 'light'
  if (preferences.theme === Theme.Dark) return 'dark'
  return electron.nativeTheme?.shouldUseDarkColors ? 'dark' : 'light'
}

export function editorPreferences(): EditorPreferences {
  current ??= load()
  return current
}

/** Uma vez na inicialização, depois de `installBundledDictionary`. */
export function applyStoredPreferences(): void {
  const active = editorPreferences()
  if (electron.session?.defaultSession) applySpellChecker(electron.session.defaultSession, active.spellcheck)
  // Antes de a janela abrir, para a primeira pintura já sair na cor certa.
  if (electron.nativeTheme) electron.nativeTheme.themeSource = active.theme
}

/** Remendo, e não o conjunto inteiro: quem clica em "marcas de formatação" não opina sobre ortografia. */
export function updatePreferences(patch: EditorPreferencesPatch): EditorPreferences {
  const active = editorPreferences()
  // Chave por chave: espalhar o remendo apagaria a chave presente com
  // `undefined`. O nome apagado volta a ser o do sistema na hora.
  const next: EditorPreferences = withAuthor({
    spellcheck: patch.spellcheck ?? active.spellcheck,
    invisibleCharacters: patch.invisibleCharacters ?? active.invisibleCharacters,
    typography: patch.typography ?? active.typography,
    language: patch.language ?? active.language,
    theme: patch.theme ?? active.theme,
    readingMode: patch.readingMode ?? active.readingMode,
    showToolbar: patch.showToolbar ?? active.showToolbar,
    showStatusBar: patch.showStatusBar ?? active.showStatusBar,
    zoom: patch.zoom ?? active.zoom,
    zoomFit: patch.zoomFit ?? active.zoomFit,
    navigationPane: patch.navigationPane ?? active.navigationPane,
    commentsPane: patch.commentsPane ?? active.commentsPane,
    authorName: patch.authorName ?? active.authorName,
  })

  const spellcheckChanged = next.spellcheck !== active.spellcheck
  const themeChanged = next.theme !== active.theme

  // Sobre o próprio tipo, para a preferência nova não escapar de uma lista à mão.
  const keys = Object.keys(next) as (keyof EditorPreferences)[]
  const unchanged = keys.every((key) => next[key] === active[key])

  // Sem isto, reabrir o menu com o mesmo valor gravaria o arquivo e redesenharia o editor.
  if (unchanged) return active

  current = next
  getStore().set('preferences', next)

  if (spellcheckChanged && electron.session?.defaultSession) {
    applySpellChecker(electron.session.defaultSession, next.spellcheck)
  }

  // O renderer resolve `system` por `matchMedia`, que enxerga o `themeSource`.
  if (themeChanged && electron.nativeTheme) electron.nativeTheme.themeSource = next.theme

  // Sempre, inclusive para quem pediu: assim a barra e o menu mostram o mesmo.
  broadcastPush(IpcChannel.PreferencesChanged, next)
  for (const listener of listeners) listener(next)

  return next
}

/** Um emissor, e não uma chamada a `refreshMenu`: o menu já importa este módulo. */
export function onPreferencesChanged(listener: (preferences: EditorPreferences) => void): void {
  listeners.add(listener)
}
