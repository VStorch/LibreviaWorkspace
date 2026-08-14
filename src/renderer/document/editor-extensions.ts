import type { Extensions } from '@tiptap/core'
import StarterKit from '@tiptap/starter-kit'
import Highlight from '@tiptap/extension-highlight'
import Image from '@tiptap/extension-image'
import TextAlign from '@tiptap/extension-text-align'
import { TableKit } from '@tiptap/extension-table'
import { CharacterCount } from '@tiptap/extensions'
import {
  BackgroundColor,
  Color,
  FontFamily,
  FontSize,
  LineHeight,
  TextStyle,
} from '@tiptap/extension-text-style'
import { Indent } from './extensions/indent.js'
import { Caps, SmallCaps } from './extensions/letter-case.js'
import { PageBreak } from './extensions/page-break.js'
import { SearchReplace, type SearchStatus } from './extensions/search-replace.js'

/**
 * Conjunto de extensões do editor.
 *
 * Cobre a seção "Texto" e "Inserção" da especificação. Três extensões são
 * nossas porque não existem oficialmente: recuo, quebra de página e
 * localizar/substituir.
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

    Highlight.configure({ multicolor: true }),
    TextAlign.configure({ types: ['heading', 'paragraph'] }),

    Image.configure({
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
    // Vieram do corpus real: `w:caps` e `w:smallCaps` aparecem 45 vezes.
    Caps,
    SmallCaps,
    PageBreak,
    SearchReplace.configure({ onStatusChange: onSearchStatusChange }),
  ]
}
