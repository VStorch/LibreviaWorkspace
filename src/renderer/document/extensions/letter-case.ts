import { Mark, mergeAttributes } from '@tiptap/core'

/**
 * Asked for by the corpus: a 15-page document uses `w:caps` and `w:smallCaps` 45 times. Two marks,
 * because in OOXML they are independent properties.
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
