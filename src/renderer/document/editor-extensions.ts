import { InputRule, type Extensions } from '@tiptap/core'
import StarterKit from '@tiptap/starter-kit'
import {
  InvisibleCharacter,
  InvisibleCharacters,
  HardBreakNode,
  ParagraphNode,
  SpaceCharacter,
} from '@tiptap/extension-invisible-characters'
import Typography from '@tiptap/extension-typography'
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
import { Indent } from './extensions/indent.js'
import { Caps, SmallCaps } from './extensions/letter-case.js'
import { PageBreak } from './extensions/page-break.js'
import { Pagination } from './extensions/pagination.js'
import { ParagraphCommands } from './extensions/paragraph-commands.js'
import { SearchReplace, type SearchStatus } from './extensions/search-replace.js'
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
      link: {
        // Links do documento não navegam dentro do aplicativo: são abertos no
        // navegador do sistema, e só depois de passarem pela allowlist de
        // esquema no processo main (ver src/main/security-policy.ts).
        openOnClick: false,
        autolink: true,
        HTMLAttributes: { rel: 'noopener noreferrer' },
      },
      // O histórico do Tiptap já responde a Ctrl+Z e Ctrl+Y.
      undoRedo: { depth: 200 },
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
      inline: false,
      // Imagens entram como data URI, validadas no processo main antes de
      // chegarem aqui. SVG é recusado lá: é vetor de script.
      allowBase64: true,
    }),

    TableKit.configure({
      table: { resizable: true, allowTableNodeSelection: true },
    }),

    // Alimenta a contagem exibida na barra de status.
    CharacterCount,

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
    // Fundo, espaçamento e entrelinha do parágrafo — no OOXML são
    // propriedades do bloco, e é o que faz `Heading1` virar barra colorida.
    BlockFormat,
    // A identidade que o bloco traz do `.docx`. Sem ela a gravação cirúrgica
    // deixa de reconhecer o que não mudou e regenera o documento inteiro.
    BlockIdentity,
    // Vieram do corpus real: `w:caps` e `w:smallCaps` aparecem 45 vezes.
    Caps,
    SmallCaps,
    PageBreak,
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
