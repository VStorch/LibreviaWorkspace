import { Mark, mergeAttributes } from '@tiptap/core'

/**
 * Maiúsculas e versalete.
 *
 * Não existem no StarterKit, e entraram porque o corpus real pediu: o documento
 * de 15 páginas usa `w:caps` e `w:smallCaps` 45 vezes (docs/01-corpus-docx.md).
 * Sem eles, esses trechos abririam sem a formatação e a gravação os devolveria
 * como texto comum — perda silenciosa, justamente o que a Fase 4 combate.
 *
 * São dois marks e não um atributo de `textStyle` porque no OOXML são duas
 * propriedades independentes, que podem estar ligadas ao mesmo tempo.
 */

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    letterCase: {
      toggleCaps: () => ReturnType
      toggleSmallCaps: () => ReturnType
    }
  }
}

export const Caps = Mark.create({
  name: 'caps',

  parseHTML() {
    return [{ style: 'text-transform=uppercase' }, { tag: 'span[data-caps]' }]
  },

  renderHTML({ HTMLAttributes }) {
    return [
      'span',
      mergeAttributes(HTMLAttributes, { 'data-caps': '', style: 'text-transform: uppercase' }),
      0,
    ]
  },

  addCommands() {
    return {
      toggleCaps:
        () =>
        ({ commands }) =>
          commands.toggleMark(this.name),
    }
  },
})

export const SmallCaps = Mark.create({
  name: 'smallCaps',

  parseHTML() {
    return [{ style: 'font-variant=small-caps' }, { tag: 'span[data-small-caps]' }]
  },

  renderHTML({ HTMLAttributes }) {
    return [
      'span',
      mergeAttributes(HTMLAttributes, { 'data-small-caps': '', style: 'font-variant: small-caps' }),
      0,
    ]
  },

  addCommands() {
    return {
      toggleSmallCaps:
        () =>
        ({ commands }) =>
          commands.toggleMark(this.name),
    }
  },
})
