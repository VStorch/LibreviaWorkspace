import { InputRule, type Attribute, type Extensions } from '@tiptap/core'
import StarterKit from '@tiptap/starter-kit'
import {
  InvisibleCharacter,
  InvisibleCharacters,
  HardBreakNode,
  ParagraphNode,
  SpaceCharacter,
} from '@tiptap/extension-invisible-characters'
import Typography from '@tiptap/extension-typography'
import Link from '@tiptap/extension-link'
import Highlight from '@tiptap/extension-highlight'
import TextAlign from '@tiptap/extension-text-align'
import { TableKit } from '@tiptap/extension-table'
import Superscript from '@tiptap/extension-superscript'
import Subscript from '@tiptap/extension-subscript'
import { CharacterCount } from '@tiptap/extensions'
import {
  BackgroundColor,
  Color,
  FontFamily,
  FontSize,
  LineHeight,
  TextStyle,
} from '@tiptap/extension-text-style'
import { BlockFormat } from './extensions/block-format.js'
import { DocumentImage } from './extensions/document-image.js'
import { BlockIdentity } from './extensions/block-identity.js'
import { BookmarkEnd, BookmarkStart, Bookmarks } from './extensions/bookmark.js'
import { Field } from './extensions/field.js'
import { TableOfContents } from './extensions/table-of-contents.js'
import { Indent } from './extensions/indent.js'
import { ListNumbering } from './extensions/list-numbering.js'
import { Caps, SmallCaps } from './extensions/letter-case.js'
import { PageBreak } from './extensions/page-break.js'
import { ReadOnlyGuard } from './extensions/read-only-guard.js'
import { Pagination } from './extensions/pagination.js'
import { ZoomedColumnResize } from './extensions/zoomed-column-resize.js'
import { ParagraphCommands } from './extensions/paragraph-commands.js'
import { CharacterStyle, StyleCommands } from './extensions/style-commands.js'
import { SearchReplace, type SearchStatus } from './extensions/search-replace.js'
import { TableLook } from './extensions/table-look.js'
import { WordShortcuts } from './extensions/word-shortcuts.js'

/** O que o editor precisa saber das preferências de edição ao ser montado. */
export interface EditorToolOptions {
  /**
   * Se a autocorreção tipográfica está ligada — consultada **a cada** regra.
   *
   * Uma função, e não um booleano: as regras de entrada são registradas quando o
   * editor nasce, e a preferência muda com ele no ar. Com um valor fixo aqui,
   * desligar a autocorreção só valeria no próximo documento aberto.
   */
  readonly isTypographyEnabled?: () => boolean
  /** Estado inicial das marcas de formatação. Depois quem manda é o comando. */
  readonly invisibleCharactersVisible?: boolean
}

/**
 * Conjunto de extensões do editor.
 *
 * Cobre a seção "Texto" e "Inserção" da especificação. Quatro extensões são
 * nossas porque não existem oficialmente: recuo, quebra de página,
 * localizar/substituir e os comandos do diálogo de parágrafo.
 */
export function buildEditorExtensions(
  onSearchStatusChange: (status: SearchStatus) => void,
  options: EditorToolOptions = {},
): Extensions {
  return [
    StarterKit.configure({
      // O link é o de baixo, com menos atributos — ver `DocumentLink`.
      link: false,
      // O histórico do Tiptap já responde a Ctrl+Z e Ctrl+Y.
      undoRedo: { depth: 200 },
      // O parágrafo vazio que o Tiptap acrescenta no fim do documento quando o
      // último bloco não é parágrafo. Depois de um título ele não serve para
      // nada — Enter no fim do título já abre um parágrafo —, e custava caro:
      // os documentos do corpus terminam num `Heading1`, e gravar sem editar
      // acrescentava um `<w:p/>` ao arquivo (um bloco reescrito, e às vezes uma
      // linha a mais no pé da última folha).
      trailingNode: { notAfter: ['paragraph', 'heading'] },
    }),

    DocumentLink.configure({
      // Links do documento não navegam dentro do aplicativo: são abertos no
      // navegador do sistema, e só depois de passarem pela allowlist de
      // esquema no processo main (ver src/main/security-policy.ts). O link para
      // um marcador (`#nome`) leva ao marcador — ver bookmark.ts.
      openOnClick: false,
      autolink: true,
      HTMLAttributes: { rel: 'noopener noreferrer' },
    }),

    // `TextStyle` é o suporte para cor, fonte, tamanho e espaçamento — todos
    // guardados como atributos de uma marca só.
    TextStyle,
    Color,
    BackgroundColor,
    FontFamily,
    FontSize,
    LineHeight,

    // Sobrescrito e subscrito. São marcas de verdade, e não um atributo de
    // `textStyle`, porque no OOXML são um `w:vertAlign` — uma propriedade só,
    // com dois valores que se excluem, e as extensões oficiais já se excluem
    // uma à outra. Enquanto não existiam, o texto sobrescrito de um documento
    // abria como texto comum e voltava assim para o arquivo.
    Superscript,
    Subscript,

    Highlight.configure({ multicolor: true }),
    TextAlign.configure({ types: ['heading', 'paragraph'] }),

    DocumentImage.configure({
      // Em linha, porque é assim que ela está no arquivo: no OOXML não existe
      // imagem fora de parágrafo, e a que o sidecar lê chega dentro do dela. Como
      // bloco, o schema não a aceitava ali; o documento abria assim mesmo — o
      // JSON não é conferido —, mas a primeira mudança de atributo partia o
      // parágrafo, a imagem descia para um novo e voltava ao arquivo como imagem
      // nova, com outro `wp:docPr` e outro relacionamento.
      inline: true,
      // Imagens entram como data URI, validadas no processo main antes de
      // chegarem aqui. SVG é recusado lá: é vetor de script.
      allowBase64: true,
    }),

    TableKit.configure({
      // `resizable` é o arrasto da divisória das colunas; a medida resultante
      // vai para o `colwidth` das células e dali para o `w:tblGrid`.
      table: { resizable: true, allowTableNodeSelection: true },
    }),
    // Borda e sombreamento de célula, e os comandos de largura de coluna que o
    // diálogo de propriedades usa. Ver table-look.ts.
    TableLook,
    // O arrasto da divisória dividido pelo zoom da folha.
    ZoomedColumnResize,

    // Alimenta a contagem exibida na barra de status.
    CharacterCount,

    // O somente leitura vale para comando, e não só para o teclado.
    ReadOnlyGuard,

    // Marcas de formatação: ¶ no fim do parágrafo, ponto no espaço, seta na
    // tabulação e ¬ na quebra de linha. São decorações, então não entram no HTML
    // que gera o PDF — o papel nunca as mostra, como no Word.
    //
    // `injectCSS: false` de propósito: o estilo que vem com a extensão desenha as
    // marcas com `line-height: 1em`, e num documento paginado ao vivo uma marca
    // que mexa na medida da linha desloca a quebra de página. O nosso mora em
    // `content-styles.ts`, com `line-height: 0` — a mesma lição que o sobrescrito
    // deixou.
    InvisibleCharacters.configure({
      visible: options.invisibleCharactersVisible ?? false,
      injectCSS: false,
      builders: [
        new SpaceCharacter(),
        // A tabulação não vem na lista padrão da extensão, e num documento de
        // escritório ela é justamente o que se procura quando o alinhamento saiu
        // errado.
        new InvisibleCharacter({ type: 'tab', predicate: (char) => char === '\t' }),
        new ParagraphNode(),
        new HardBreakNode(),
      ],
    }),

    // Autocorreção tipográfica: aspas curvas, travessão, reticências. Ver
    // `guardedTypography` para o que foi desligado e por quê.
    guardedTypography(options.isTypographyEnabled ?? (() => true)),

    Indent,
    // Os comandos que o diálogo de parágrafo usa: escrevem o formulário inteiro
    // numa transação só, para que desfazer não peça oito `Ctrl+Z`.
    ParagraphCommands,
    // Aplicar estilo, limpar a formatação direta, o "desligado" que vence o
    // estilo e o Enter que passa ao estilo seguinte. Ver style-commands.ts.
    StyleCommands,
    CharacterStyle,
    // Fundo, espaçamento e entrelinha do parágrafo — no OOXML são
    // propriedades do bloco, e é o que faz `Heading1` virar barra colorida.
    BlockFormat,
    // A numeração das listas contada como o Word conta: por definição, por
    // nível e com o formato de cada nível. Ver list-numbering.ts.
    ListNumbering,
    // A identidade que o bloco traz do `.docx`. Sem ela a gravação cirúrgica
    // deixa de reconhecer o que não mudou e regenera o documento inteiro.
    BlockIdentity,
    // Vieram do corpus real: `w:caps` e `w:smallCaps` aparecem 45 vezes.
    Caps,
    SmallCaps,
    PageBreak,
    // Referências (M8): as duas pontas de cada marcador, e os comandos, o clique
    // no link interno e a colagem sem marcador repetido. Ver bookmark.ts.
    BookmarkStart,
    BookmarkEnd,
    Bookmarks,
    // Os campos (PAGEREF, REF, SEQ…) como nós com instrução e resultado, e o
    // sumário como bloco. Ver field.ts e table-of-contents.ts.
    Field,
    TableOfContents,
    // Guarda os vãos entre as folhas. Quem os calcula é `usePagination`; aqui
    // fica só o lugar onde eles vivem, para acompanharem a edição sem que o
    // documento saiba que existem.
    Pagination,
    SearchReplace.configure({ onStatusChange: onSearchStatusChange }),
    // Por último na lista e com prioridade alta no próprio arquivo: é ele que
    // decide `Ctrl+E`, disputado com a marca de código. Ver word-shortcuts.ts.
    WordShortcuts,
  ]
}

/**
 * A autocorreção tipográfica, com interruptor e com duas regras a menos.
 *
 * ## O que ficou de fora
 *
 * `<-` e `->` viravam flechas, e `3 x 4` virava `3 × 4`. As duas são regras que o
 * Word **não** tem em português, e as duas atrapalham texto técnico: ninguém que
 * escreve `a -> b` num procedimento quis uma flecha, e `x` entre números aparece
 * em dimensão de imagem e em código. O resto do conjunto é o que o Word faz:
 * aspas curvas, travessão, reticências, `(c)`, `(r)`, `(tm)`, frações e expoentes.
 *
 * ## Por que embrulhar em vez de configurar
 *
 * As regras de entrada do Tiptap são registradas quando o editor nasce, e não há
 * como retirar as de uma extensão depois. O embrulho consulta a preferência no
 * momento em que a regra ia disparar: devolver `null` ali é dizer "não houve
 * correção", e o texto digitado fica como está.
 *
 * ## E para desfazer uma correção só
 *
 * `Backspace` logo depois da substituição desfaz **só ela** — `—silent` volta a
 * `--silent` —, como no Word. Não há atalho nosso para isso: quem trata a tecla é
 * o `Keymap` do próprio Tiptap, que tenta `undoInputRule` antes de qualquer outra
 * coisa. O embrulho preserva o `undoable` de cada regra justamente para que essa
 * tecla continue funcionando; sem ele, desligar a autocorreção seria a única saída
 * para quem escreve texto técnico.
 */
function guardedTypography(enabled: () => boolean): Extensions[number] {
  return Typography.configure({
    leftArrow: false,
    rightArrow: false,
    multiplication: false,
  }).extend({
    addInputRules() {
      return (this.parent?.() ?? []).map(
        (rule) =>
          new InputRule({
            find: rule.find,
            handler: (props) => (enabled() ? rule.handler(props) : null),
            // Sem isto o `Backspace` não desfaria a substituição: é esta bandeira
            // que faz o plugin guardar o que desfazer. Ver o comentário acima.
            undoable: rule.undoable,
          }),
      )
    },
  })
}

/**
 * O link com os atributos que o `.docx` tem: o endereço e a dica.
 *
 * O padrão da extensão materializa `target`, `rel` e `class` em **toda** marca, e
 * o `w:hyperlink` não tem nenhum deles: o nó que voltava do editor trazia dois
 * atributos que o sidecar não lera, a impressão digital divergia e todo parágrafo
 * com link era reescrito ao salvar sem ninguém tê-lo tocado. Os dois continuam no
 * HTML — quem os escreve é `HTMLAttributes`, que não passa pelo modelo.
 */
const DocumentLink = Link.extend({
  addAttributes() {
    const inherited: Record<string, Attribute> = this.parent?.() ?? {}
    return { href: inherited['href'] ?? {}, title: inherited['title'] ?? {} }
  },
})
