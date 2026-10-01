import { describe, expect, it } from 'vitest'
import type { DocumentComment } from './model.js'
import {
  commentAnchorIdsOfJson,
  commentsOutsideOf,
  initialsOf,
  nextCommentId,
  paragraphsOfText,
  resolveComments,
} from './comments.js'

const comment = (id: string, parentId?: string): DocumentComment => ({
  id,
  author: 'Ana',
  date: '',
  paragraphs: [id],
  done: false,
  ...(parentId === undefined ? {} : { parentId }),
})

const library = [comment('0'), comment('1', '0'), comment('2'), comment('3', '2'), comment('4')]

describe('resolveComments', () => {
  it('a conversa vale enquanto as pontas dela estão no texto, e as respostas com ela', () => {
    expect(resolveComments(new Set(['0', '4']), library).map((c) => c.id)).toEqual(['0', '1', '4'])
  })

  it('a conversa ancorada fora do corpo vale sem ponta no texto', () => {
    expect(resolveComments(new Set(), library, new Set(['2'])).map((c) => c.id)).toEqual(['2', '3'])
  })

  it('a resposta cujo comentário não está na biblioteca é conversa própria', () => {
    const orphan = [comment('9', '8')]
    expect(resolveComments(new Set(['9']), orphan)).toHaveLength(1)
    expect(resolveComments(new Set(), orphan)).toHaveLength(0)
  })
})

describe('commentsOutsideOf', () => {
  it('as conversas sem ponta no documento aberto, e só as que abrem conversa', () => {
    const doc = {
      type: 'doc',
      content: [
        {
          type: 'paragraph',
          content: [
            { type: 'commentStart', attrs: { cid: '0' } },
            { type: 'text', text: 'a' },
            { type: 'commentEnd', attrs: { cid: '0' } },
            { type: 'commentEnd', attrs: { cid: '4' } },
          ],
        },
      ],
    }
    expect([...commentAnchorIdsOfJson(doc)].sort()).toEqual(['0', '4'])
    expect(commentsOutsideOf(doc, library)).toEqual(['2'])
  })
})

describe('auxiliares', () => {
  it('o id novo é o próximo depois do maior da biblioteca', () => {
    expect(nextCommentId([])).toBe('0')
    expect(nextCommentId([comment('3'), comment('10'), comment('x')])).toBe('11')
  })

  it('iniciais e parágrafos', () => {
    expect(initialsOf('vinícius storch')).toBe('VS')
    expect(initialsOf('  ')).toBe('')
    expect(paragraphsOfText('um\r\ndois')).toEqual(['um', 'dois'])
  })
})
