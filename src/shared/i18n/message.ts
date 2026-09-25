import type { Language } from './language.js'

/**
 * O texto de uma entrada num idioma.
 *
 * Ou é uma frase, ou é um par singular/plural. O par não é enfeite: a barra de
 * status diz "1 página" e "2 páginas", e uma frase só com `{count}` no meio
 * obrigaria cada chamada a resolver o plural por fora — que é exatamente onde
 * ele para de ser resolvido.
 */
export type Message = string | PluralMessage

export interface PluralMessage {
  readonly one: string
  readonly other: string
}

/**
 * Uma entrada do catálogo: a mesma frase nos dois idiomas, **lado a lado**.
 *
 * Um arquivo por idioma seria o arranjo comum, e é o arranjo que apodrece: a
 * chave nova entra num deles, ninguém percebe a falta no outro, e a tela abre
 * metade traduzida. Aqui a falta é erro de compilação — `Entry` exige as duas
 * chaves — e quem revisa vê as duas frases na mesma linha, que é a única forma
 * de notar que elas deixaram de dizer a mesma coisa.
 */
export interface Entry {
  readonly pt: Message
  readonly en: Message
}

/** O formato de um catálogo. Usado com `satisfies`, para as chaves ficarem literais. */
export type Catalog = Readonly<Record<string, Entry>>

/**
 * Os valores que uma frase interpola.
 *
 * `count` é especial: além de ser interpolado como qualquer outro, é ele que
 * escolhe entre singular e plural.
 */
export type Vars = Readonly<Record<string, string | number>>

/**
 * Resolve uma entrada: escolhe o idioma, escolhe o número, preenche os buracos.
 *
 * Função pura e sem estado — é o que permite testá-la sem subir nem Electron
 * nem React, e é o que permite o main e o renderer usarem a mesma.
 */
export function format(entry: Entry, language: Language, vars?: Vars): string {
  const message = entry[language]
  const count = vars?.['count']

  const text =
    typeof message === 'string'
      ? message
      : // Português e inglês concordam na regra que nos importa: um é singular,
        // todo o resto é plural — inclusive o zero ("0 páginas", "0 pages").
        // `Intl.PluralRules` saberia mais, e não saberia nada de útil a mais
        // para estes dois idiomas.
        count === 1
        ? message.one
        : message.other

  if (vars === undefined) return text
  return interpolate(text, vars)
}

/**
 * Troca `{nome}` pelo valor.
 *
 * Uma varredura só, e não uma substituição por chave: com uma por chave, um
 * valor que por acaso contivesse `{outra}` seria interpolado na volta seguinte.
 * Nome de arquivo faz isso — e é justamente nome de arquivo que mais entra
 * nestas frases.
 *
 * Buraco sem valor fica como está. Apagá-lo esconderia o erro; deixá-lo visível
 * faz a falta aparecer na tela, onde alguém a conserta.
 */
function interpolate(text: string, vars: Vars): string {
  return text.replace(/\{(\w+)\}/g, (whole, name: string) => {
    const value = vars[name]
    return value === undefined ? whole : String(value)
  })
}
