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
import { CommentEnd, CommentStart, Comments } from './extensions/comment.js'
import { NoteRef } from './extensions/note-ref.js'
import type { DocumentNotes } from '@services/document/model.js'
import { ZeroWidthAnchors } from './extensions/zero-width.js'
import { CountWithoutDeletions, TrackChanges } from './extensions/track-changes.js'
import { TrackInput } from './extensions/track-input.js'
import { RevisionViewExtension } from './extensions/revision-view.js'
import { Field } from './extensions/field.js'
import { MathEditing, MathNode } from './extensions/math.js'
import { TableOfContents } from './extensions/table-of-contents.js'
import { Indent } from './extensions/indent.js'
import { ListNumbering } from './extensions/list-numbering.js'
import { Caps, SmallCaps } from './extensions/letter-case.js'
import { PageBreak } from './extensions/page-break.js'
import { ReadOnlyGuard } from './extensions/read-only-guard.js'
import { Pagination } from './extensions/pagination.js'
import { SectionGeometry, SectionMarks } from './extensions/section-geometry.js'
import { ZoomedColumnResize } from './extensions/zoomed-column-resize.js'
import { ParagraphCommands } from './extensions/paragraph-commands.js'
import { CharacterStyle, StyleCommands } from './extensions/style-commands.js'
import { SearchReplace, type SearchStatus } from './extensions/search-replace.js'
import { TableLook } from './extensions/table-look.js'
import { WordShortcuts } from './extensions/word-shortcuts.js'

export interface EditorToolOptions {
  /** Uma função, consultada a cada regra: as regras nascem com o editor, e a preferência muda com ele no ar. */
  readonly isTypographyEnabled?: () => boolean
  /** Só o estado inicial; depois quem manda é o comando. */
  readonly invisibleCharactersVisible?: boolean
  /** Se a conversa está na biblioteca de comentários — ver `withoutCommentAnchors`. */
  readonly isKnownComment?: (cid: string) => boolean
  /** Consultado a cada transação. */
  readonly isTrackingChanges?: () => boolean
  readonly revisionAuthor?: () => string
  readonly notes?: () => DocumentNotes | undefined
}

export function buildEditorExtensions(
  onSearchStatusChange: (status: SearchStatus) => void,
  options: EditorToolOptions = {},
): Extensions {
  return [
    ...contentExtensions(),

    // A contagem sem o texto excluído por uma revisão, como no Word.
    CountWithoutDeletions,

    // O somente leitura vale para comando, e não só para o teclado.
    ReadOnlyGuard,

    // `injectCSS: false`: o estilo da extensão usa `line-height: 1em`, e uma marca
    // que mude a medida da linha desloca a quebra de página (ver `content-styles.ts`).
    InvisibleCharacters.configure({
      visible: options.invisibleCharactersVisible ?? false,
      injectCSS: false,
      builders: [
        new SpaceCharacter(),
        // A tabulação, que a extensão não traz e é o que se procura quando o alinhamento saiu errado.
        new InvisibleCharacter({ type: 'tab', predicate: (char) => char === '\t' }),
        new ParagraphNode(),
        new HardBreakNode(),
      ],
    }),

    guardedTypography(options.isTypographyEnabled ?? (() => true)),

    Indent,
    // O formulário inteiro numa transação, para desfazer não pedir oito `Ctrl+Z`.
    ParagraphCommands,
    StyleCommands,
    CharacterStyle,
    BlockFormat,
    ListNumbering,
    // Sem a identidade do bloco, a gravação cirúrgica regeneraria o documento inteiro.
    BlockIdentity,
    Caps,
    SmallCaps,
    PageBreak,
    BookmarkStart,
    BookmarkEnd,
    Bookmarks,
    CommentStart,
    CommentEnd,
    Comments.configure({ isKnown: options.isKnownComment }),
    NoteRef.configure({ notes: options.notes }),
    ...TrackChanges,
    // Antes do controle do que se digita: o Backspace passa pelo escondido antes.
    RevisionViewExtension,
    TrackInput.configure({
      isTracking: options.isTrackingChanges ?? (() => false),
      author: options.revisionAuthor ?? (() => ''),
    }),
    ZeroWidthAnchors,
    Field,
    TableOfContents,
    MathNode,
    MathEditing,
    // Só guarda os vãos que `usePagination` calcula, para acompanharem a edição.
    Pagination,
    SectionGeometry,
    SectionMarks,
    SearchReplace.configure({ onStatusChange: onSearchStatusChange }),
    // Prioridade alta: decide `Ctrl+E`, disputado com a marca de código.
    WordShortcuts,
  ]
}

/** O texto, a formatação, a imagem e a tabela; o resto é do editor de documento. */
function contentExtensions(): Extensions {
  return [
    StarterKit.configure({
      link: false,
      undoRedo: { depth: 200 },
      // Sem o parágrafo vazio depois do título: o corpus termina em `Heading1`, e
      // gravar sem editar acrescentaria um `<w:p/>`.
      trailingNode: { notAfter: ['paragraph', 'heading'] },
    }),

    DocumentLink.configure({
      // Os links abrem no navegador do sistema, depois da lista de esquemas do main.
      openOnClick: false,
      autolink: true,
      HTMLAttributes: { rel: 'noopener noreferrer' },
    }),

    TextStyle,
    Color,
    BackgroundColor,
    FontFamily,
    FontSize,
    LineHeight,

    // Marcas, e não atributo de `textStyle`: no OOXML são um `w:vertAlign`, de valores que se excluem.
    Superscript,
    Subscript,

    Highlight.configure({ multicolor: true }),
    TextAlign.configure({ types: ['heading', 'paragraph'] }),

    DocumentImage.configure({
      // Em linha, como no arquivo: no OOXML não há imagem fora de parágrafo, e
      // como bloco a primeira mudança de atributo partiria o parágrafo.
      inline: true,
      // Data URI, validado no main, que recusa SVG.
      allowBase64: true,
    }),

    TableKit.configure({
      table: { resizable: true, allowTableNodeSelection: true },
    }),
    TableLook,
    ZoomedColumnResize,
  ]
}

/**
 * Sem `<-`, `->` e `3 x 4`, que o Word não tem em português e atrapalham texto
 * técnico. Embrulhada, e não configurada, porque as regras do Tiptap nascem com
 * o editor: devolver `null` é "não houve correção". O `undoable` é preservado
 * para o Backspace logo depois desfazer só a substituição, como no Word.
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
            // É o que faz o Backspace desfazer a substituição.
            undoable: rule.undoable,
          }),
      )
    },
  })
}

/**
 * O padrão da extensão poria `target`, `rel` e `class` em toda marca, que o
 * `w:hyperlink` não tem: a impressão digital divergiria e todo parágrafo com link
 * seria reescrito. No HTML eles continuam, por `HTMLAttributes`.
 */
const DocumentLink = Link.extend({
  addAttributes() {
    const inherited: Record<string, Attribute> = this.parent?.() ?? {}
    return { href: inherited['href'] ?? {}, title: inherited['title'] ?? {} }
  },
})
