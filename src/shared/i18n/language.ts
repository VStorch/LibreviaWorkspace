export const Language = {
  Portuguese: 'pt',
  English: 'en',
} as const

export type Language = (typeof Language)[keyof typeof Language]

export const LANGUAGES: readonly Language[] = [Language.Portuguese, Language.English]

/** Each language name written in that language, as every language list does. */
export const LANGUAGE_NAMES: Readonly<Record<Language, string>> = {
  pt: 'Português',
  en: 'English',
}

export function languageFromLocale(locale: string): Language {
  return locale.toLowerCase().startsWith('pt') ? Language.Portuguese : Language.English
}
