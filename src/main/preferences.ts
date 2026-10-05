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
 * Main owns them because spelling is a `session` setting. Every change goes through
 * `updatePreferences`, which saves, applies and notifies.
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
      // A corrupt JSON does not stop the app from opening: preferences are a convenience.
      clearInvalidConfig: true,
      ...(process.versions.electron ? {} : { cwd: process.cwd() }),
    })
  }
  return storeInstance
}

/** In memory; the file is the durable copy. */
let current: EditorPreferences | null = null

const listeners = new Set<(preferences: EditorPreferences) => void>()

function load(): EditorPreferences {
  return withAuthor(loadStored())
}

/** Without an author, the system user, as Word does the first time. */
function withAuthor(preferences: EditorPreferences): EditorPreferences {
  if (preferences.authorName.trim() !== '') return preferences
  let username = ''
  try {
    username = userInfo().username
  } catch {
    // Without a system account (container), comments go out without an author.
  }
  return { ...preferences, authorName: username }
}

function loadStored(): EditorPreferences {
  if (!process.versions.electron) return DEFAULT_EDITOR_PREFERENCES

  try {
    const store = getStore()
    const stored = store.get('preferences')

    // Through the schema: the `default`s open a profile saved without these keys.
    const parsed = editorPreferencesSchema.safeParse(stored)
    const preferences = parsed.success ? parsed.data : DEFAULT_EDITOR_PREFERENCES

    // On first run the language comes from the operating system. The question is "was the key
    // saved?", and the `default` would erase that distinction, so it is read from the raw object.
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

/** `nativeTheme.shouldUseDarkColors` changes by itself when the user switches the system theme. */
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

/** Once at startup, after `installBundledDictionary`. */
export function applyStoredPreferences(): void {
  const active = editorPreferences()
  if (electron.session?.defaultSession) applySpellChecker(electron.session.defaultSession, active.spellcheck)
  // Before the window opens, so the first paint has the right colors.
  if (electron.nativeTheme) electron.nativeTheme.themeSource = active.theme
}

/** A patch, not the whole set: whoever toggles formatting marks has no opinion about spelling. */
export function updatePreferences(patch: EditorPreferencesPatch): EditorPreferences {
  const active = editorPreferences()
  const next = mergedPreferences(active, patch)

  const spellcheckChanged = next.spellcheck !== active.spellcheck
  const themeChanged = next.theme !== active.theme

  // From the type itself, so a new preference cannot escape a hand-written list.
  const keys = Object.keys(next) as (keyof EditorPreferences)[]
  const unchanged = keys.every((key) => next[key] === active[key])

  // Otherwise reopening the menu with the same value would save the file and redraw the editor.
  if (unchanged) return active

  current = next
  getStore().set('preferences', next)

  if (spellcheckChanged && electron.session?.defaultSession) {
    applySpellChecker(electron.session.defaultSession, next.spellcheck)
  }

  // The renderer resolves `system` through `matchMedia`, which sees `themeSource`.
  if (themeChanged && electron.nativeTheme) electron.nativeTheme.themeSource = next.theme

  // Always, including to the sender: the toolbar and the menu show the same thing.
  broadcastPush(IpcChannel.PreferencesChanged, next)
  for (const listener of listeners) listener(next)

  return next
}

function mergedPreferences(active: EditorPreferences, patch: EditorPreferencesPatch): EditorPreferences {
  // Key by key: spreading the patch would erase a key present with `undefined`. A cleared name
  // falls back to the system user immediately.
  return withAuthor({
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
}

/** An emitter, not a call to `refreshMenu`: the menu already imports this module. */
export function onPreferencesChanged(listener: (preferences: EditorPreferences) => void): void {
  listeners.add(listener)
}
