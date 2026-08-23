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

          /**
           * A fonte do próprio bloco.
           *
           * Não é redundante com a marca do texto: **a altura da linha nasce da
           * fonte do elemento**, não do que está escrito dentro dele. Um
           * parágrafo de 10 pt num bloco que o CSS declara com 12 pt continua
           * ocupando 12 pt de altura, e um título de 10 pt vira uma barra alta
           * demais porque o editor desenha títulos grandes.
           */
          fontFamily: {
            default: null,
            parseHTML: (element) => element.style.fontFamily || null,
            renderHTML: (attributes) => {
              const family = attributes['fontFamily']
              return typeof family === 'string' && family.length > 0
                ? { style: `font-family: ${family}` }
                : {}
            },
          },

          fontSize: {
            default: null,
            parseHTML: (element) => element.style.fontSize || null,
            renderHTML: (attributes) => {
              const size = attributes['fontSize']
              return typeof size === 'string' && size.length > 0 ? { style: `font-size: ${size}` } : {}
            },
          },

          /**
           * A entrelinha, já como o CSS a escreve.
           *
           * Texto, e não número: o espaçamento simples do Word — o que o
           * arquivo quer dizer quando não diz nada — é a altura que a própria
           * fonte pede, e em CSS isso se chama `normal`. Nenhum fator o imita,
           * e enquanto o padrão do editor valia para o documento importado cada
           * linha saía meia altura mais arejada do que no Word.
           */
          lineHeight: {
            default: null,
            parseHTML: (element) => element.style.lineHeight || null,
            renderHTML: (attributes) => {
              const value = attributes['lineHeight']
              if (typeof value === 'number') {
                return Number.isFinite(value) && value > 0 ? { style: `line-height: ${value}` } : {}
              }
              return typeof value === 'string' && value.length > 0 ? { style: `line-height: ${value}` } : {}
            },
          },

          /**
           * "Manter com o próximo" do Word (`w:keepNext`). Não muda a
           * aparência: diz que este bloco não pode ficar sozinho no pé da
           * página, e é o que a marca de fim de página e a exportação usam.
           */
          /**
           * A folha termina depois deste bloco.
           *
           * Vem de um `w:br w:type="page"` que o Word gravou **dentro** do
           * parágrafo. Emiti-lo como nó ali dentro poria um bloco em posição de
           * linha — inválido no schema, e o serializador o desalojaria ao
           * atravessar HTML, desalinhando os índices entre tela e papel.
           */
          breakAfter: {
            default: null,
            parseHTML: (element) => (element.hasAttribute('data-break-after') ? true : null),
            renderHTML: (attributes) => (attributes['breakAfter'] === true ? { 'data-break-after': '' } : {}),
          },

          keepNext: {
            default: null,
            parseHTML: (element) => element.hasAttribute('data-keep-next') || null,
            renderHTML: (attributes) => (attributes['keepNext'] === true ? { 'data-keep-next': '' } : {}),
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
