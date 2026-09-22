/**
 * Catálogo de caracteres especiais.
 *
 * É o conjunto que o Word deixa à mão na aba "Símbolo" — o que se procura de
 * verdade num documento: aspas tipográficas, travessão, moeda,
 * matemática, letra grega e marca registrada. Uma tabela Unicode completa seria
 * mais completa e menos útil: quem precisa de ❡ sabe achá-lo.
 *
 * Cada caractere leva nome traduzido porque o nome é o que o leitor de tela
 * anuncia — um botão chamado "—" não diz nada a quem não o vê.
 */

import { Language, translate, type MessageKey } from '@shared/i18n/index.js'

export interface SpecialCharacter {
  /** O caractere em si, do jeito que entra no documento. */
  readonly char: string
  /** Chave de tradução do nome. */
  readonly nameKey: MessageKey
  /** Como ele se chama. Vira o nome acessível do botão e a dica do mouse. */
  readonly name: string
}

export interface SpecialCharacterGroup {
  readonly labelKey: MessageKey
  readonly label: string
  readonly characters: readonly SpecialCharacter[]
}

interface CharDef {
  readonly char: string
  readonly nameKey: MessageKey
}

interface GroupDef {
  readonly labelKey: MessageKey
  readonly characters: readonly CharDef[]
}

const GROUPS_DEF: readonly GroupDef[] = [
  {
    labelKey: 'chars.group.punctuation',
    characters: [
      { char: '“', nameKey: 'chars.punct.doubleQuoteOpen' },
      { char: '”', nameKey: 'chars.punct.doubleQuoteClose' },
      { char: '‘', nameKey: 'chars.punct.singleQuoteOpen' },
      { char: '’', nameKey: 'chars.punct.singleQuoteClose' },
      { char: '«', nameKey: 'chars.punct.angleQuoteOpen' },
      { char: '»', nameKey: 'chars.punct.angleQuoteClose' },
      { char: '—', nameKey: 'chars.punct.emDash' },
      { char: '–', nameKey: 'chars.punct.enDash' },
      { char: '…', nameKey: 'chars.punct.ellipsis' },
      { char: '·', nameKey: 'chars.punct.middleDot' },
      { char: '•', nameKey: 'chars.punct.bullet' },
      { char: '§', nameKey: 'chars.punct.section' },
      { char: '¶', nameKey: 'chars.punct.pilcrow' },
      { char: '†', nameKey: 'chars.punct.dagger' },
      { char: '‡', nameKey: 'chars.punct.doubleDagger' },
      { char: '‰', nameKey: 'chars.punct.perMille' },
      { char: '¿', nameKey: 'chars.punct.invertedQuestion' },
      { char: '¡', nameKey: 'chars.punct.invertedExclamation' },
      // Espaço inquebrável: é ele que impede "R$" de ficar no fim de uma linha e
      // o valor na seguinte.
      { char: '\u00a0', nameKey: 'chars.punct.nonBreakingSpace' },
    ],
  },
  {
    labelKey: 'chars.group.currency',
    characters: [
      { char: '$', nameKey: 'chars.currency.dollar' },
      { char: '€', nameKey: 'chars.currency.euro' },
      { char: '£', nameKey: 'chars.currency.pound' },
      { char: '¥', nameKey: 'chars.currency.yen' },
      { char: '¢', nameKey: 'chars.currency.cent' },
      { char: '₽', nameKey: 'chars.currency.ruble' },
      { char: '₹', nameKey: 'chars.currency.rupee' },
      { char: '¤', nameKey: 'chars.currency.generic' },
    ],
  },
  {
    labelKey: 'chars.group.math',
    characters: [
      { char: '×', nameKey: 'chars.math.multiplication' },
      { char: '÷', nameKey: 'chars.math.division' },
      { char: '±', nameKey: 'chars.math.plusMinus' },
      { char: '≠', nameKey: 'chars.math.notEqual' },
      { char: '≈', nameKey: 'chars.math.approxEqual' },
      { char: '≤', nameKey: 'chars.math.lessOrEqual' },
      { char: '≥', nameKey: 'chars.math.greaterOrEqual' },
      { char: '∞', nameKey: 'chars.math.infinity' },
      { char: '√', nameKey: 'chars.math.squareRoot' },
      { char: '∑', nameKey: 'chars.math.summation' },
      { char: '∏', nameKey: 'chars.math.product' },
      { char: '∫', nameKey: 'chars.math.integral' },
      { char: '∂', nameKey: 'chars.math.partialDifferential' },
      { char: '∆', nameKey: 'chars.math.delta' },
      { char: '°', nameKey: 'chars.math.degree' },
      { char: '′', nameKey: 'chars.math.prime' },
      { char: '″', nameKey: 'chars.math.doublePrime' },
      { char: '½', nameKey: 'chars.math.half' },
      { char: '¼', nameKey: 'chars.math.quarter' },
      { char: '¾', nameKey: 'chars.math.threeQuarters' },
      { char: '¹', nameKey: 'chars.math.superscriptOne' },
      { char: '²', nameKey: 'chars.math.superscriptTwo' },
      { char: '³', nameKey: 'chars.math.superscriptThree' },
      { char: 'µ', nameKey: 'chars.math.micro' },
    ],
  },
  {
    labelKey: 'chars.group.greek',
    characters: [
      { char: 'α', nameKey: 'chars.greek.alpha' },
      { char: 'β', nameKey: 'chars.greek.beta' },
      { char: 'γ', nameKey: 'chars.greek.gamma' },
      { char: 'δ', nameKey: 'chars.greek.delta' },
      { char: 'ε', nameKey: 'chars.greek.epsilon' },
      { char: 'θ', nameKey: 'chars.greek.theta' },
      { char: 'λ', nameKey: 'chars.greek.lambda' },
      { char: 'μ', nameKey: 'chars.greek.mu' },
      { char: 'π', nameKey: 'chars.greek.pi' },
      { char: 'ρ', nameKey: 'chars.greek.rho' },
      { char: 'σ', nameKey: 'chars.greek.sigma' },
      { char: 'τ', nameKey: 'chars.greek.tau' },
      { char: 'φ', nameKey: 'chars.greek.phi' },
      { char: 'χ', nameKey: 'chars.greek.chi' },
      { char: 'ψ', nameKey: 'chars.greek.psi' },
      { char: 'ω', nameKey: 'chars.greek.omega' },
      { char: 'Γ', nameKey: 'chars.greek.capitalGamma' },
      { char: 'Δ', nameKey: 'chars.greek.capitalDelta' },
      { char: 'Θ', nameKey: 'chars.greek.capitalTheta' },
      { char: 'Λ', nameKey: 'chars.greek.capitalLambda' },
      { char: 'Π', nameKey: 'chars.greek.capitalPi' },
      { char: 'Σ', nameKey: 'chars.greek.capitalSigma' },
      { char: 'Φ', nameKey: 'chars.greek.capitalPhi' },
      { char: 'Ω', nameKey: 'chars.greek.capitalOmega' },
    ],
  },
  {
    labelKey: 'chars.group.marksArrows',
    characters: [
      { char: '©', nameKey: 'chars.marks.copyright' },
      { char: '®', nameKey: 'chars.marks.registered' },
      { char: '™', nameKey: 'chars.marks.trademark' },
      { char: '℠', nameKey: 'chars.marks.serviceMark' },
      { char: 'ª', nameKey: 'chars.marks.feminineOrdinal' },
      { char: 'º', nameKey: 'chars.marks.masculineOrdinal' },
      { char: '№', nameKey: 'chars.marks.numero' },
      { char: '✓', nameKey: 'chars.marks.check' },
      { char: '✗', nameKey: 'chars.marks.cross' },
      { char: '★', nameKey: 'chars.marks.blackStar' },
      { char: '☆', nameKey: 'chars.marks.whiteStar' },
      { char: '→', nameKey: 'chars.arrows.right' },
      { char: '←', nameKey: 'chars.arrows.left' },
      { char: '↑', nameKey: 'chars.arrows.up' },
      { char: '↓', nameKey: 'chars.arrows.down' },
    ],
  },
]

export function specialCharacterGroups(language: Language = Language.Portuguese): readonly SpecialCharacterGroup[] {
  return GROUPS_DEF.map((group) => ({
    labelKey: group.labelKey,
    label: translate(language, group.labelKey),
    characters: group.characters.map((character) => ({
      char: character.char,
      nameKey: character.nameKey,
      name: translate(language, character.nameKey),
    })),
  }))
}

export const SPECIAL_CHARACTER_GROUPS: readonly SpecialCharacterGroup[] = specialCharacterGroups()

/** Todos os caracteres do catálogo, na ordem em que aparecem. */
export function allSpecialCharacters(language: Language = Language.Portuguese): readonly SpecialCharacter[] {
  return specialCharacterGroups(language).flatMap((group) => group.characters)
}
