export const Language = {
  Portuguese: 'pt',
  English: 'en',
} as const

export type Language = (typeof Language)[keyof typeof Language]

export const LANGUAGES: readonly Language[] = [Language.Portuguese, Language.English]

/** O nome de cada idioma, escrito nele mesmo — como toda lista de idiomas faz. */
export const LANGUAGE_NAMES: Readonly<Record<Language, string>> = {
  pt: 'Português',
  en: 'English',
}

/** Qualquer localidade que não seja português cai em inglês. */
export function languageFromLocale(locale: string): Language {
  return locale.toLowerCase().startsWith('pt') ? Language.Portuguese : Language.English
}
