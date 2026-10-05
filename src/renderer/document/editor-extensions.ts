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
  /**
   * A function, queried on each rule: the rules are created with the editor, and the preference
   * changes while it is running.
   */
  readonly isTypographyEnabled?: () => boolean
  /** Only the initial state; afterwards the command rules. */
  readonly invisibleCharactersVisible?: boolean
  /** Whether the thread is in the comment library; see `withoutCommentAnchors`. */
  readonly isKnownComment?: (cid: string) => boolean
  /** Queried on each transaction. */
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

    // Counting without text deleted by a revision, as in Word.
    CountWithoutDeletions,

    // Read-only applies to commands, not just to the keyboard.
    ReadOnlyGuard,

    // `injectCSS: false`: the extension style uses `line-height: 1em`, and a mark that changes the
    // line height shifts the page break (see `content-styles.ts`).
    InvisibleCharacters.configure({
      visible: options.invisibleCharactersVisible ?? false,
      injectCSS: false,
      builders: [
        new SpaceCharacter(),
        // The tab, which the extension lacks and is what one looks for when alignment came out
        // wrong.
        new InvisibleCharacter({ type: 'tab', predicate: (char) => char === '\t' }),
        new ParagraphNode(),
        new HardBreakNode(),
      ],
    }),

    guardedTypography(options.isTypographyEnabled ?? (() => true)),

    Indent,
    // The whole form in one transaction, so undo does not take eight `Ctrl+Z`.
    ParagraphCommands,
    StyleCommands,
    CharacterStyle,
    BlockFormat,
    ListNumbering,
    // Without block identity, the surgical save would regenerate the whole document.
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
    // Before input tracking: Backspace goes through the hidden text first.
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
    // Only keeps the gaps `usePagination` computes, so they follow editing.
    Pagination,
    SectionGeometry,
    SectionMarks,
    SearchReplace.configure({ onStatusChange: onSearchStatusChange }),
    // High priority: it decides `Ctrl+E`, contested by the code mark.
    WordShortcuts,
  ]
}

/** Text, formatting, images and tables; the rest belongs to the document editor. */
function contentExtensions(): Extensions {
  return [
    StarterKit.configure({
      link: false,
      undoRedo: { depth: 200 },
      // No empty paragraph after a heading: the corpus ends in `Heading1`, and saving without
      // editing would add a `<w:p/>`.
      trailingNode: { notAfter: ['paragraph', 'heading'] },
    }),

    DocumentLink.configure({
      // Links open in the system browser, after main's scheme allowlist.
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

    // Marks, not a `textStyle` attribute: in OOXML they are one `w:vertAlign`, with mutually
    // exclusive values.
    Superscript,
    Subscript,

    Highlight.configure({ multicolor: true }),
    TextAlign.configure({ types: ['heading', 'paragraph'] }),

    DocumentImage.configure({
      // Inline, as in the file: OOXML has no image outside a paragraph, and as a block the first
      // attribute change would split the paragraph.
      inline: true,
      // A data URI validated in main, which refuses SVG.
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
 * Without `<-`, `->` and `3 x 4`, which Word lacks in Portuguese and which get in the way of
 * technical text. Wrapped, not configured, because Tiptap's rules are created with the editor:
 * returning `null` means "no correction happened". `undoable` is kept so Backspace right after
 * undoes only the replacement, as in Word.
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
            // That is what makes Backspace undo the replacement.
            undoable: rule.undoable,
          }),
      )
    },
  })
}

/**
 * The extension default would put `target`, `rel` and `class` on every mark, which `w:hyperlink`
 * lacks: the fingerprint would diverge and every paragraph with a link would be rewritten. In HTML
 * they remain, through `HTMLAttributes`.
 */
const DocumentLink = Link.extend({
  addAttributes() {
    const inherited: Record<string, Attribute> = this.parent?.() ?? {}
    return { href: inherited['href'] ?? {}, title: inherited['title'] ?? {} }
  },
})
