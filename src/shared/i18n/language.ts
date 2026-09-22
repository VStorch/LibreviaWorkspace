/**
 * Os idiomas em que o aplicativo fala.
 *
 * Dois, e não uma lista aberta: cada idioma novo é um valor a mais em **toda**
 * entrada do catálogo, e o compilador cobra os que faltam. É de propósito —
 * um idioma pela metade é pior que idioma nenhum, porque a tela fica metade em
 * cada língua e ninguém sabe se é bug ou tradução pendente.
 *
 * Mora em `shared` porque as duas pontas precisam: o menu nativo é montado no
 * main, e o resto da interface no renderer. Uma cópia em cada lado divergiria
 * no primeiro rótulo mudado às pressas.
 */
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

/**
 * O idioma que uma etiqueta de localidade pede.
 *
 * Recebe o que o sistema diz — `pt-BR`, `en-US`, `pt`, `es-AR` — e responde com
 * um dos dois que existem. Qualquer coisa que não seja português cai em inglês,
 * que é a escolha menos ruim para quem não fala nenhum dos dois: o inglês é o
 * idioma que mais gente lê por acidente.
 *
 * Só o prefixo é olhado. `pt-PT` e `pt-BR` são o mesmo catálogo aqui; separá-los
 * seria prometer uma distinção que as traduções não fazem.
 */
export function languageFromLocale(locale: string): Language {
  return locale.toLowerCase().startsWith('pt') ? Language.Portuguese : Language.English
}
