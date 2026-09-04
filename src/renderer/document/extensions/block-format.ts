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
    // A lista entra junto: no arquivo ela não existe como bloco — são
    // parágrafos numerados —, mas na árvore do editor é um elemento de verdade,
    // e sem espaçamento declarado ele recebe o do editor.
    return { types: ['paragraph', 'heading', 'bulletList', 'orderedList'] }
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

          /**
           * A marca da lista, como o documento a declara.
           *
           * Sem ela o CSS escolhe a bolinha, e o documento pede um quadrado —
           * `w:lvlText` guarda o caractere, e ele costuma vir da área de uso
           * privado do Unicode, que é como o Word grava os glifos das fontes
           * Symbol e Wingdings. Quem os traduz é o leitor.
           */
          marker: {
            default: null,
            parseHTML: (element) => element.getAttribute('data-marker'),
            renderHTML: (attributes) => {
              const value = attributes['marker']
              if (typeof value !== 'string' || value.length === 0) return {}
              // Como custom property, e não `list-style-type`: o marcador é
              // desenhado por um `::before` do item, que é o único jeito de
              // controlar a distância dele até o texto — o recuo pendente do
              // Word. Aspas simples porque o valor entra numa string de CSS, e
              // uma aspa dentro dele a fecharia.
              return {
                'data-marker': value,
                style: `--marca: '${value.replace(/['\\]/g, '\\$&')}'`,
              }
            },
          },

          /**
           * Onde o texto do item começa, em milímetros.
           *
           * `w:ind/@left` do nível. Sem ele o item sai colado na margem, e não
           * no recuo que o documento pede.
           */
          indentMm: {
            default: null,
            parseHTML: (element) => element.getAttribute('data-indent-mm'),
            renderHTML: (attributes) => {
              const value = Number(attributes['indentMm'])
              // A medida sai duas vezes: como recuo de verdade e como variável.
              // A imagem ancorada não é texto — no Word ela se posiciona pela
              // coluna, e não pelo recuo do parágrafo —, e é pela variável que
              // ela desconta o recuo de volta.
              return Number.isFinite(value) && value > 0
                ? {
                    'data-indent-mm': String(value),
                    style: `padding-left: ${value}mm; --recuo: ${value}mm`,
                  }
                : {}
            },
          },

          /**
           * Recuo da direita, em milímetros.
           *
           * `w:ind/@right`. Estreita a coluna do parágrafo, e por isso muda
           * onde a linha quebra — não é decoração.
           */
          indentRightMm: {
            default: null,
            parseHTML: (element) => element.getAttribute('data-indent-right-mm'),
            renderHTML: (attributes) => {
              const value = Number(attributes['indentRightMm'])
              return Number.isFinite(value) && value > 0
                ? {
                    'data-indent-right-mm': String(value),
                    style: `padding-right: ${value}mm; --recuo-direita: ${value}mm`,
                  }
                : {}
            },
          },

          /**
           * Onde a primeira linha começa, em milímetros, a contar do recuo.
           *
           * Positivo é `w:firstLine` e negativo é `w:hanging` — o Word os grava
           * como dois atributos, mas eles são a mesma medida com o sinal
           * trocado, e é assim que o CSS a entende.
           */
          firstLineMm: {
            default: null,
            parseHTML: (element) => element.getAttribute('data-first-line-mm'),
            renderHTML: (attributes) => {
              const value = Number(attributes['firstLineMm'])
              return Number.isFinite(value) && value !== 0
                ? { 'data-first-line-mm': String(value), style: `text-indent: ${value}mm` }
                : {}
            },
          },

          /**
           * Quanto o marcador fica antes do texto do item.
           *
           * `w:ind/@hanging` do nível — o recuo pendente do Word. Sem ele a
           * marca sai encostada na primeira letra.
           */
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

          /**
           * O parágrafo é só a marca de uma seção.
           *
           * No OOXML a seção termina num `w:sectPr` guardado dentro do `w:pPr`
           * de um parágrafo vazio: o parágrafo **é** a marca, e o LibreOffice
           * não lhe dá altura nenhuma. Sem isto cada marca valia uma linha, e o
           * documento de evidências do corpus tem seis delas espalhadas pelo
           * meio do texto.
           */
          sectionMark: {
            default: null,
            parseHTML: (element) => element.hasAttribute('data-section-mark') || null,
            renderHTML: (attributes) =>
              attributes['sectionMark'] === true ? { 'data-section-mark': '' } : {},
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
