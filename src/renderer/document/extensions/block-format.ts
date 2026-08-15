import { Extension } from '@tiptap/core'

/**
 * Formatação de bloco vinda do documento: fundo, espaçamento e entrelinha.
 *
 * Existe porque o Tiptap trata isso como estilo de texto, e no OOXML é
 * propriedade do **parágrafo**. A diferença não é acadêmica: no corpus real, o
 * estilo `Heading1` é uma barra vermelha com texto branco — o fundo é do
 * parágrafo, e sem ele o título vira texto solto no meio da página.
 *
 * Os valores chegam já resolvidos pelo sidecar (padrões do documento + estilo +
 * formatação direta), porque o editor não tem noção de estilo.
 */

export interface BlockFormatOptions {
  types: string[]
}

// Zero é valor legítimo — "sem espaço antes" é instrução do documento, e
// descartá-lo deixaria a margem padrão do editor reaparecer.
const pointsToCss = (value: unknown): string | null => {
  if (value === null || value === undefined) return null
  const points = Number(value)
  return Number.isFinite(points) && points >= 0 ? `${points}pt` : null
}

export const BlockFormat = Extension.create<BlockFormatOptions>({
  name: 'blockFormat',

  addOptions() {
    return { types: ['paragraph', 'heading'] }
  },

  addGlobalAttributes() {
    return [
      {
        types: this.options.types,
        attributes: {
          background: {
            default: null,
            parseHTML: (element) => element.style.backgroundColor || null,
            renderHTML: (attributes) => {
              const color = attributes['background']
              return typeof color === 'string' && color.length > 0
                ? { style: `background-color: ${color}` }
                : {}
            },
          },

          spaceBefore: {
            default: null,
            parseHTML: (element) => element.style.marginTop || null,
            renderHTML: (attributes) => {
              const value = pointsToCss(attributes['spaceBefore'])
              return value === null ? {} : { style: `margin-top: ${value}` }
            },
          },

          spaceAfter: {
            default: null,
            parseHTML: (element) => element.style.marginBottom || null,
            renderHTML: (attributes) => {
              const value = pointsToCss(attributes['spaceAfter'])
              return value === null ? {} : { style: `margin-bottom: ${value}` }
            },
          },

          lineHeight: {
            default: null,
            parseHTML: (element) => element.style.lineHeight || null,
            renderHTML: (attributes) => {
              const factor = Number(attributes['lineHeight'])
              return Number.isFinite(factor) && factor > 0 ? { style: `line-height: ${factor}` } : {}
            },
          },

          /**
           * O identificador do estilo do Word. Não muda nada na tela: viaja
           * junto para que um parágrafo editado continue apontando o estilo
           * original na hora de gravar.
           */
          styleId: {
            default: null,
            parseHTML: (element) => element.getAttribute('data-style-id'),
            renderHTML: (attributes) => {
              const id = attributes['styleId']
              return typeof id === 'string' && id.length > 0 ? { 'data-style-id': id } : {}
            },
          },
        },
      },
    ]
  },
})
