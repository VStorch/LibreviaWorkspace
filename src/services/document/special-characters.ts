/**
 * Catálogo de caracteres especiais.
 *
 * É o conjunto que o Word deixa à mão na aba "Símbolo" — o que se procura de
 * verdade num documento em português: aspas tipográficas, travessão, moeda,
 * matemática, letra grega e marca registrada. Uma tabela Unicode completa seria
 * mais completa e menos útil: quem precisa de ❡ sabe achá-lo.
 *
 * Cada caractere leva nome em português porque o nome é o que o leitor de tela
 * anuncia — um botão chamado "—" não diz nada a quem não o vê.
 */

export interface SpecialCharacter {
  /** O caractere em si, do jeito que entra no documento. */
  readonly char: string
  /** Como ele se chama. Vira o nome acessível do botão e a dica do mouse. */
  readonly name: string
}

export interface SpecialCharacterGroup {
  readonly label: string
  readonly characters: readonly SpecialCharacter[]
}

export const SPECIAL_CHARACTER_GROUPS: readonly SpecialCharacterGroup[] = [
  {
    label: 'Pontuação',
    characters: [
      { char: '“', name: 'abre aspas duplas' },
      { char: '”', name: 'fecha aspas duplas' },
      { char: '‘', name: 'abre aspas simples' },
      { char: '’', name: 'fecha aspas simples' },
      { char: '«', name: 'abre aspas angulares' },
      { char: '»', name: 'fecha aspas angulares' },
      { char: '—', name: 'travessão' },
      { char: '–', name: 'meia-risca' },
      { char: '…', name: 'reticências' },
      { char: '·', name: 'ponto mediano' },
      { char: '•', name: 'marcador' },
      { char: '§', name: 'parágrafo de lei' },
      { char: '¶', name: 'marca de parágrafo' },
      { char: '†', name: 'obelisco' },
      { char: '‡', name: 'obelisco duplo' },
      { char: '‰', name: 'por milhar' },
      { char: '¿', name: 'abre interrogação' },
      { char: '¡', name: 'abre exclamação' },
      // Espaço inquebrável: é ele que impede "R$" de ficar no fim de uma linha e
      // o valor na seguinte.
      { char: '\u00a0', name: 'espaço inquebrável' },
    ],
  },
  {
    label: 'Moeda',
    characters: [
      { char: '$', name: 'dólar' },
      { char: '€', name: 'euro' },
      { char: '£', name: 'libra' },
      { char: '¥', name: 'iene' },
      { char: '¢', name: 'centavo' },
      { char: '₽', name: 'rublo' },
      { char: '₹', name: 'rupia' },
      { char: '¤', name: 'moeda genérica' },
    ],
  },
  {
    label: 'Matemática',
    characters: [
      { char: '×', name: 'multiplicação' },
      { char: '÷', name: 'divisão' },
      { char: '±', name: 'mais ou menos' },
      { char: '≠', name: 'diferente' },
      { char: '≈', name: 'aproximadamente' },
      { char: '≤', name: 'menor ou igual' },
      { char: '≥', name: 'maior ou igual' },
      { char: '∞', name: 'infinito' },
      { char: '√', name: 'raiz quadrada' },
      { char: '∑', name: 'somatório' },
      { char: '∏', name: 'produtório' },
      { char: '∫', name: 'integral' },
      { char: '∂', name: 'derivada parcial' },
      { char: '∆', name: 'variação' },
      { char: '°', name: 'grau' },
      { char: '′', name: 'minuto' },
      { char: '″', name: 'segundo' },
      { char: '½', name: 'um meio' },
      { char: '¼', name: 'um quarto' },
      { char: '¾', name: 'três quartos' },
      { char: '¹', name: 'expoente um' },
      { char: '²', name: 'expoente dois' },
      { char: '³', name: 'expoente três' },
      { char: 'µ', name: 'micro' },
    ],
  },
  {
    label: 'Grego',
    characters: [
      { char: 'α', name: 'alfa' },
      { char: 'β', name: 'beta' },
      { char: 'γ', name: 'gama' },
      { char: 'δ', name: 'delta' },
      { char: 'ε', name: 'épsilon' },
      { char: 'θ', name: 'teta' },
      { char: 'λ', name: 'lambda' },
      { char: 'μ', name: 'mi' },
      { char: 'π', name: 'pi' },
      { char: 'ρ', name: 'rô' },
      { char: 'σ', name: 'sigma' },
      { char: 'τ', name: 'tau' },
      { char: 'φ', name: 'fi' },
      { char: 'χ', name: 'qui' },
      { char: 'ψ', name: 'psi' },
      { char: 'ω', name: 'ômega' },
      { char: 'Γ', name: 'gama maiúsculo' },
      { char: 'Δ', name: 'delta maiúsculo' },
      { char: 'Θ', name: 'teta maiúsculo' },
      { char: 'Λ', name: 'lambda maiúsculo' },
      { char: 'Π', name: 'pi maiúsculo' },
      { char: 'Σ', name: 'sigma maiúsculo' },
      { char: 'Φ', name: 'fi maiúsculo' },
      { char: 'Ω', name: 'ômega maiúsculo' },
    ],
  },
  {
    label: 'Marcas e setas',
    characters: [
      { char: '©', name: 'direito autoral' },
      { char: '®', name: 'marca registrada' },
      { char: '™', name: 'marca comercial' },
      { char: '℠', name: 'marca de serviço' },
      { char: 'ª', name: 'ordinal feminino' },
      { char: 'º', name: 'ordinal masculino' },
      { char: '№', name: 'número' },
      { char: '✓', name: 'certo' },
      { char: '✗', name: 'errado' },
      { char: '★', name: 'estrela cheia' },
      { char: '☆', name: 'estrela vazia' },
      { char: '→', name: 'seta à direita' },
      { char: '←', name: 'seta à esquerda' },
      { char: '↑', name: 'seta para cima' },
      { char: '↓', name: 'seta para baixo' },
    ],
  },
]

/** Todos os caracteres do catálogo, na ordem em que aparecem. */
export function allSpecialCharacters(): readonly SpecialCharacter[] {
  return SPECIAL_CHARACTER_GROUPS.flatMap((group) => group.characters)
}
