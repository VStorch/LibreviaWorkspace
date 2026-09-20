import type { Extensions } from '@tiptap/core'
import StarterKit from '@tiptap/starter-kit'
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

/**
 * Conjunto de extensões do editor.
 *
 * Cobre a seção "Texto" e "Inserção" da especificação. Quatro extensões são
 * nossas porque não existem oficialmente: recuo, quebra de página,
 * localizar/substituir e os comandos do diálogo de parágrafo.
 */
export function buildEditorExtensions(onSearchStatusChange: (status: SearchStatus) => void): Extensions {
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
