import { Extension, type Attributes } from '@tiptap/core'

/**
 * No OOXML fundo, espaçamento e entrelinha são do **parágrafo**, e não estilo de
 * texto como no Tiptap: no corpus, o `Heading1` é uma barra vermelha com texto
 * branco. No parágrafo do corpo só vem a formatação direta, e o inline vence a
 * regra do estilo; na lista, na célula e no rascunho antigo os valores chegam
 * já resolvidos, porque ali nenhuma regra de estilo alcança.
 */

export interface BlockFormatOptions {
  types: string[]
}

// Zero é "sem espaço antes", uma instrução do documento.
const pointsToCss = (value: unknown): string | null => {
  if (value === null || value === undefined) return null
  const points = Number(value)
  return Number.isFinite(points) && points >= 0 ? `${points}pt` : null
}

/** O zero conta: é o parágrafo desfazendo o recuo do estilo. `Number(null)` é zero, daí a conferência antes. */
const declaredMeasure = (value: unknown): number | null => {
  if (value === null || value === undefined || value === '') return null
  const measure = Number(value)
  return Number.isFinite(measure) ? measure : null
}

const BLOCK_ATTRIBUTES: Attributes = {
  background: {
    default: null,
    parseHTML: (element) => element.style.backgroundColor || null,
    renderHTML: (attributes) => {
      const color = attributes['background']
      return typeof color === 'string' && color.length > 0 ? { style: `background-color: ${color}` } : {}
    },
  },

  /** `w:lvlText`, muitas vezes na área de uso privado, como o Word grava Symbol e Wingdings. */
  marker: {
    default: null,
    parseHTML: (element) => element.getAttribute('data-marker'),
    renderHTML: (attributes) => {
      const value = attributes['marker']
      if (typeof value !== 'string' || value.length === 0) return {}
      // Variável, e não `list-style-type`: o `::before` do item é o único jeito de
      // controlar a distância até o texto. Aspas simples, porque o valor entra numa string de CSS.
      return {
        'data-marker': value,
        style: `--marca: '${value.replace(/['\\]/g, '\\$&')}'`,
      }
    },
  },

  /** `w:ind/@left` do nível. */
  indentMm: {
    default: null,
    parseHTML: (element) => element.getAttribute('data-indent-mm'),
    renderHTML: (attributes) => {
      const value = declaredMeasure(attributes['indentMm'])
      // Também como variável: a imagem ancorada se posiciona pela coluna, e desconta o recuo por ela.
      return value !== null && value >= 0
        ? {
            'data-indent-mm': String(value),
            style: `padding-left: ${value}mm; --recuo: ${value}mm`,
          }
        : {}
    },
  },

  /** `w:ind/@right`: estreita a coluna e muda onde a linha quebra. */
  indentRightMm: {
    default: null,
    parseHTML: (element) => element.getAttribute('data-indent-right-mm'),
    renderHTML: (attributes) => {
      const value = declaredMeasure(attributes['indentRightMm'])
      return value !== null && value >= 0
        ? {
            'data-indent-right-mm': String(value),
            style: `padding-right: ${value}mm; --recuo-direita: ${value}mm`,
          }
        : {}
    },
  },

  /** Positivo é `w:firstLine` e negativo é `w:hanging`, a mesma medida com o sinal trocado. */
  firstLineMm: {
    default: null,
    parseHTML: (element) => element.getAttribute('data-first-line-mm'),
    renderHTML: (attributes) => {
      const value = declaredMeasure(attributes['firstLineMm'])
      return value !== null ? { 'data-first-line-mm': String(value), style: `text-indent: ${value}mm` } : {}
    },
  },

  /** `w:ind/@hanging` do nível: a distância do marcador até o texto. */
  hangingMm: {
    default: null,
    parseHTML: (element) => element.getAttribute('data-hanging-mm'),
    renderHTML: (attributes) => {
      const value = Number(attributes['hangingMm'])
      return Number.isFinite(value) && value > 0
        ? { 'data-hanging-mm': String(value), style: `--pendente: ${value}mm` }
        : {}
    },
  },

  /** O parágrafo vazio que guarda o `w:sectPr` **é** a marca, e o LibreOffice não lhe dá altura. */
  sectionMark: {
    default: null,
    parseHTML: (element) => element.hasAttribute('data-section-mark') || null,
    renderHTML: (attributes) => (attributes['sectionMark'] === true ? { 'data-section-mark': '' } : {}),
  },

  /**
   * O id da seção em `sections` que termina aqui, como o `w:sectPr` no OOXML.
   * Não passa adiante no Enter: a marca repetida faria duas seções com o mesmo id.
   */
  sectionBreak: {
    default: null,
    keepOnSplit: false,
    parseHTML: (element) => element.getAttribute('data-section-break'),
    renderHTML: (attributes) =>
      typeof attributes['sectionBreak'] === 'string'
        ? { 'data-section-break': attributes['sectionBreak'] }
        : {},
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

  /** A altura da linha nasce da fonte do elemento, e não da do texto dentro dele. */
  fontFamily: {
    default: null,
    parseHTML: (element) => element.style.fontFamily || null,
    renderHTML: (attributes) => {
      const family = attributes['fontFamily']
      return typeof family === 'string' && family.length > 0 ? { style: `font-family: ${family}` } : {}
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

  /** O simples do Word é a altura que a fonte pede, que no CSS é `normal`: nenhum fator o imita. */
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
   * Um `w:br w:type="page"` gravado **dentro** do parágrafo: como nó, ficaria em posição de
   * linha, inválido no schema, e desalinharia os índices entre tela e papel.
   */
  breakAfter: {
    default: null,
    parseHTML: (element) => (element.hasAttribute('data-break-after') ? true : null),
    renderHTML: (attributes) => (attributes['breakAfter'] === true ? { 'data-break-after': '' } : {}),
  },

  /** `w:br w:type="column"`, pelo mesmo motivo de `breakAfter`; não passa adiante no Enter. */
  columnBreakAfter: {
    default: null,
    keepOnSplit: false,
    parseHTML: (element) => (element.hasAttribute('data-column-break') ? true : null),
    renderHTML: (attributes) => (attributes['columnBreakAfter'] === true ? { 'data-column-break': '' } : {}),
  },

  /** `w:keepNext`: não muda a aparência; a marca de fim de página e a exportação o usam. */
  keepNext: {
    default: null,
    parseHTML: (element) => element.hasAttribute('data-keep-next') || null,
    renderHTML: (attributes) => (attributes['keepNext'] === true ? { 'data-keep-next': '' } : {}),
  },

  /** `w:keepLines`. `false` é o parágrafo desfazendo o que o estilo liga. */
  keepLines: {
    default: null,
    parseHTML: (element) => element.hasAttribute('data-keep-lines') || null,
    renderHTML: (attributes) => (attributes['keepLines'] === true ? { 'data-keep-lines': '' } : {}),
  },

  /** `w:widowControl`: ausente é ligado, como no Word. */
  widowControl: {
    default: null,
    parseHTML: (element) => (element.hasAttribute('data-widows-allowed') ? false : null),
    renderHTML: (attributes) => (attributes['widowControl'] === false ? { 'data-widows-allowed': '' } : {}),
  },

  /** Não muda a tela: o parágrafo editado continua apontando o estilo original ao gravar. */
  styleId: {
    default: null,
    parseHTML: (element) => element.getAttribute('data-style-id'),
    renderHTML: (attributes) => {
      const id = attributes['styleId']
      return typeof id === 'string' && id.length > 0 ? { 'data-style-id': id } : {}
    },
  },
}

const LIST_ATTRIBUTES: Attributes = {
  /** `w:numId`: sem ele o gravador escreveria `w:numId w:val="0"`, que no OOXML é "sem numeração". */
  numId: {
    default: null,
    parseHTML: (element) => {
      const value = Number(element.getAttribute('data-num-id'))
      return Number.isInteger(value) && value > 0 ? value : null
    },
    renderHTML: (attributes) => {
      const value = Number(attributes['numId'])
      return Number.isInteger(value) && value > 0 ? { 'data-num-id': String(value) } : {}
    },
  },
}

export const BlockFormat = Extension.create<BlockFormatOptions>({
  name: 'blockFormat',

  addOptions() {
    // A lista não existe como bloco no arquivo, mas na árvore do editor existe,
    // e sem espaçamento declarado receberia o do editor.
    return { types: ['paragraph', 'heading', 'bulletList', 'orderedList'] }
  },

  addGlobalAttributes() {
    return [
      { types: this.options.types, attributes: BLOCK_ATTRIBUTES },
      // Só nas listas, as únicas com numeração.
      { types: ['bulletList', 'orderedList'], attributes: LIST_ATTRIBUTES },
    ]
  },
})
