import { Mark, mergeAttributes } from '@tiptap/core'

/**
 * Pedidos pelo corpus: um documento de 15 páginas usa `w:caps` e `w:smallCaps` 45
 * vezes. Dois marks, porque no OOXML são propriedades independentes.
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
