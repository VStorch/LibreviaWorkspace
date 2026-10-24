import { describe, expect, it } from 'vitest'
import { INVOCABLE_IPC_CHANNELS, IpcChannel, PUSH_IPC_CHANNELS } from './ipc-channels.js'
import { MAX_TEXT_LENGTH, ipcContracts, pushContracts } from './ipc.js'
import {
  MAX_AI_CHAT_MESSAGE_LENGTH,
  MAX_AI_CHAT_MESSAGES,
  MAX_AI_DOCUMENT_LENGTH,
  MAX_AI_PROPOSALS,
  MAX_AI_RESPONSE_LENGTH,
} from './limits.js'

describe('IPC contracts', () => {
  it('defines a contract for each invocable channel', () => {
    // A channel without a schema would pass unvalidated payloads to the handler.
    for (const channel of INVOCABLE_IPC_CHANNELS) {
      expect(ipcContracts[channel]).toBeDefined()
    }
  })

  it('exposes no contract for an undeclared channel', () => {
    expect(Object.keys(ipcContracts).sort()).toEqual([...INVOCABLE_IPC_CHANNELS].sort())
  })

  it('keeps the menu channel out of the invocable ones: it goes from main to renderer', () => {
    expect(INVOCABLE_IPC_CHANNELS).not.toContain(IpcChannel.MenuCommand)
    expect(ipcContracts).not.toHaveProperty(IpcChannel.MenuCommand)
  })
})

describe('main → renderer contracts', () => {
  it('defines a contract for each pushed channel', () => {
    // Without a schema nobody would check the shape, and the renderer would drop the message
    // silently, the costliest defect here.
    for (const channel of PUSH_IPC_CHANNELS) {
      expect(pushContracts[channel]).toBeDefined()
    }
  })

  it('the two directions do not mix', () => {
    for (const channel of PUSH_IPC_CHANNELS) {
      expect(INVOCABLE_IPC_CHANNELS).not.toContain(channel)
      expect(ipcContracts).not.toHaveProperty(channel)
    }
  })
})

describe('context menu target validation', () => {
  const schema = pushContracts[IpcChannel.ContextMenuRequested]

  const alvo = {
    x: 120,
    y: 40,
    editable: true,
    misspelledWord: 'abacaxxi',
    dictionarySuggestions: ['abacaxi'],
    canCut: true,
    canCopy: true,
    canPaste: true,
  }

  it('accepts what Chromium sends', () => {
    expect(schema.safeParse(alvo).success).toBe(true)
  })

  it('refuses a negative coordinate and an absurd suggestion list', () => {
    // This becomes a screen position and menu items: a negative coordinate puts the menu outside
    // the window, and thirty suggestions overflow it.
    expect(schema.safeParse({ ...alvo, x: -1 }).success).toBe(false)
    expect(
      schema.safeParse({ ...alvo, dictionarySuggestions: Array.from({ length: 30 }, () => 'x') }).success,
    ).toBe(false)
  })
})

describe('prefs:set validation', () => {
  const schema = ipcContracts[IpcChannel.PreferencesSet].request

  it('accepts a patch with a single key', () => {
    // Whoever toggles formatting marks has no opinion about spelling.
    expect(schema.parse({ invisibleCharacters: true })).toEqual({ invisibleCharacters: true })
  })

  it('refuses a non-boolean value', () => {
    expect(schema.safeParse({ spellcheck: 'sim' }).success).toBe(false)
  })
})

describe('file:save validation', () => {
  const schema = ipcContracts[IpcChannel.FileSave].request

  it('accepts a well-formed save', () => {
    expect(schema.safeParse({ path: '/home/ana/ata.txt', content: 'texto', origin: null }).success).toBe(true)
  })

  it('accepts the origin of the open document', () => {
    const payload = { path: '/home/ana/ata.docx', content: '{}', origin: '/home/ana/ata.sdoc' }
    expect(schema.safeParse(payload).success).toBe(true)
  })

  it.each([
    ['without path', { content: 'texto' }],
    ['empty path', { path: '', content: 'texto' }],
    ['without content', { path: '/a/b.txt' }],
    ['non-text path', { path: 42, content: 'texto', origin: null }],
    ['without origin', { path: '/a/b.docx', content: 'texto' }],
    ['empty origin', { path: '/a/b.docx', content: 'texto', origin: '' }],
    ['null', null],
  ])('refuses %s', (_label, payload) => {
    expect(schema.safeParse(payload).success).toBe(false)
  })

  it('refuses content above the memory ceiling', () => {
    const oversized = { path: '/a/b.txt', content: 'x'.repeat(MAX_TEXT_LENGTH + 1) }
    expect(schema.safeParse(oversized).success).toBe(false)
  })

  it('drops fields not in the contract', () => {
    const parsed = schema.parse({ path: '/a/b.txt', content: 'oi', origin: null, extra: 'ignorar' })
    expect(parsed).toEqual({ path: '/a/b.txt', content: 'oi', origin: null })
  })
})

describe('Word templates (M11)', () => {
  const open = ipcContracts[IpcChannel.TemplateOpen].request
  const opened = ipcContracts[IpcChannel.TemplateOpen].response
  const list = ipcContracts[IpcChannel.TemplateList].response

  it('opens by the source and id pair', () => {
    expect(open.safeParse({ source: 'builtin', id: 'letter.dotx' }).success).toBe(true)
    expect(open.safeParse({ source: 'user', id: '/home/ana/Modelos/proposta.dotx' }).success).toBe(true)
  })

  it.each([
    ['unknown source', { source: 'web', id: 'letter.dotx' }],
    ['empty id', { source: 'builtin', id: '' }],
    ['without id', { source: 'user' }],
  ])('refuses %s', (_label, payload) => {
    expect(open.safeParse(payload).success).toBe(false)
  })

  it('the document opened from a template carries the template mark', () => {
    const file = {
      path: '/r/letter.dotx',
      name: 'Carta.docx',
      kind: 'document',
      content: '{}',
      template: true,
    }
    expect(opened.parse({ file }).file.template).toBe(true)
  })

  it('the gallery lists both sources and the user folder', () => {
    const entry = { source: 'builtin', id: 'letter.dotx', name: 'Carta', description: 'Uma carta' }
    expect(list.safeParse({ builtin: [entry], user: [], folder: '/home/ana/Modelos' }).success).toBe(true)
    expect(list.safeParse({ builtin: [{ ...entry, source: 'outra' }], user: [], folder: '' }).success).toBe(
      false,
    )
  })
})

describe('file:open-recent validation', () => {
  const schema = ipcContracts[IpcChannel.FileOpenRecent].request

  it('requires a non-empty path', () => {
    expect(schema.safeParse({ path: '' }).success).toBe(false)
    expect(schema.safeParse({ path: '/a/b.txt' }).success).toBe(true)
  })
})

describe('fonts:list validation', () => {
  // `src/main/ipc/registry.ts` runs this schema on the real response, and `registry.test.ts` proves
  // it: a schema nobody runs is a compile-time type, not protection.
  const schema = ipcContracts[IpcChannel.FontsList].response

  it('accepts the empty list', () => {
    // A system without `fontconfig` returns nothing, and that is not an error: the toolbar keeps
    // the bundled fonts.
    expect(schema.safeParse({ families: [] }).success).toBe(true)
  })

  it('refuses an empty name and an absurd list', () => {
    // This is the output of a system program, data and not truth: an empty name would become an
    // invisible option, and a list of a hundred thousand entries would freeze the toolbar.
    expect(schema.safeParse({ families: [''] }).success).toBe(false)
    expect(schema.safeParse({ families: Array.from({ length: 4001 }, () => 'Arial') }).success).toBe(false)
  })
})

describe('dialog:confirm-discard validation', () => {
  const response = ipcContracts[IpcChannel.DialogConfirmDiscard].response

  it.each(['save', 'discard', 'cancel'])('accepts the %s answer', (choice) => {
    expect(response.safeParse({ choice }).success).toBe(true)
  })

  it('refuses an answer outside the expected three', () => {
    expect(response.safeParse({ choice: 'talvez' }).success).toBe(false)
  })
})

describe('AI assistant contracts', () => {
  const generate = ipcContracts[IpcChannel.AiGenerate].request
  const valid = {
    requestId: '0b8f3c1e-5d2a-4c7b-9e6f-1a2b3c4d5e6f',
    providerId: 'p1',
    model: 'llama3',
    action: 'rewrite',
    text: 'Um parágrafo.',
  }

  it('accepts a complete request', () => {
    expect(generate.safeParse(valid).success).toBe(true)
  })

  it('refuses an empty selection, one too long or without a valid identifier', () => {
    expect(generate.safeParse({ ...valid, text: '' }).success).toBe(false)
    expect(generate.safeParse({ ...valid, text: 'x'.repeat(20_001) }).success).toBe(false)
    expect(generate.safeParse({ ...valid, requestId: 'abc' }).success).toBe(false)
    expect(generate.safeParse({ ...valid, action: 'delete-everything' }).success).toBe(false)
  })

  describe('spreadsheet', () => {
    const sheet = ipcContracts[IpcChannel.AiSheet].request
    const formula = {
      requestId: valid.requestId,
      providerId: 'p1',
      model: '',
      task: 'formula',
      language: 'pt',
      instruction: 'some a coluna B',
      context: 'Sheet: Vendas',
    }

    it('accepts each task with what it needs', () => {
      expect(sheet.safeParse(formula).success).toBe(true)
      expect(sheet.safeParse({ ...formula, task: 'explain', instruction: '', formula: '=A1' }).success).toBe(
        true,
      )
      expect(sheet.safeParse({ ...formula, task: 'fill', rowCount: 500 }).success).toBe(true)
    })

    it('refuses a task without what it needs, too many lines and too long a context', () => {
      expect(sheet.safeParse({ ...formula, task: 'explain' }).success).toBe(false)
      expect(sheet.safeParse({ ...formula, instruction: '  ' }).success).toBe(false)
      expect(sheet.safeParse({ ...formula, task: 'fill' }).success).toBe(false)
      expect(sheet.safeParse({ ...formula, task: 'fill', rowCount: 501 }).success).toBe(false)
      expect(sheet.safeParse({ ...formula, task: 'sort' }).success).toBe(false)
      expect(sheet.safeParse({ ...formula, context: 'x'.repeat(MAX_AI_DOCUMENT_LENGTH + 1) }).success).toBe(
        false,
      )
      expect(sheet.safeParse({ ...formula, instruction: 'x'.repeat(2_001) }).success).toBe(false)
    })
  })

  describe('conversation', () => {
    const chat = ipcContracts[IpcChannel.AiChat].request
    const conversation = {
      requestId: valid.requestId,
      providerId: 'p1',
      model: '',
      messages: [
        { role: 'user', text: 'Qual o tema?' },
        { role: 'assistant', text: 'Chuva.' },
        { role: 'user', text: 'E o tom?' },
      ],
    }

    it('accepts a conversation that ends with the person', () => {
      expect(chat.safeParse(conversation).success).toBe(true)
      expect(chat.parse(conversation).proposals).toBe(false)
      expect(chat.safeParse({ ...conversation, proposals: 'sim' }).success).toBe(false)
      expect(chat.safeParse({ ...conversation, before: 'Doc', selection: 'trecho' }).success).toBe(true)
    })

    it('refuses an invalid role, an empty conversation or one that ends with the assistant', () => {
      expect(chat.safeParse({ ...conversation, messages: [{ role: 'system', text: 'x' }] }).success).toBe(
        false,
      )
      expect(chat.safeParse({ ...conversation, messages: [] }).success).toBe(false)
      expect(chat.safeParse({ ...conversation, messages: conversation.messages.slice(0, 2) }).success).toBe(
        false,
      )
    })

    it('refuses too many messages or too long a text', () => {
      const many = Array.from({ length: MAX_AI_CHAT_MESSAGES + 1 }, () => ({ role: 'user', text: 'oi' }))
      expect(chat.safeParse({ ...conversation, messages: many }).success).toBe(false)
      const long = [{ role: 'user', text: 'x'.repeat(MAX_AI_CHAT_MESSAGE_LENGTH + 1) }]
      expect(chat.safeParse({ ...conversation, messages: long }).success).toBe(false)
      const longReply = [
        { role: 'user', text: 'oi' },
        { role: 'assistant', text: 'x'.repeat(MAX_AI_RESPONSE_LENGTH + 1) },
        { role: 'user', text: 'e?' },
      ]
      expect(chat.safeParse({ ...conversation, messages: longReply }).success).toBe(false)
    })

    it('accepts numbered paragraphs only within the document size', () => {
      expect(chat.parse(conversation).paragraphs).toEqual([])
      expect(chat.safeParse({ ...conversation, paragraphs: ['Um.', 'Dois.'] }).success).toBe(true)
      const half = 'x'.repeat(MAX_AI_DOCUMENT_LENGTH / 2 + 1)
      expect(chat.safeParse({ ...conversation, paragraphs: [half, half] }).success).toBe(false)
    })
  })

  describe('proposals in the conversation reply', () => {
    const response = ipcContracts[IpcChannel.AiChat].response
    const anchor = { paragraphText: 'O gato e a chuva.' }
    const edit = {
      ...anchor,
      id: '1',
      kind: 'edit',
      paragraph: 3,
      find: 'gato',
      replace: 'cão',
      reason: 'tom',
    }
    const comment = { ...anchor, id: '2', kind: 'comment', paragraph: 1, find: 'chuva', comment: 'Fonte?' }

    it('accepts a change and a comment', () => {
      expect(response.safeParse({ status: 'done', proposals: [edit, comment] }).success).toBe(true)
      expect(response.safeParse({ status: 'done' }).success).toBe(true)
    })

    it('refuses a proposal without its text, paragraph zero, a long excerpt or too many proposals', () => {
      const noReplace = { ...anchor, id: '1', kind: 'edit', paragraph: 3, find: 'gato' }
      const noParagraphText = { ...edit, paragraphText: undefined }
      expect(response.safeParse({ status: 'done', proposals: [noParagraphText] }).success).toBe(false)
      expect(response.safeParse({ status: 'done', proposals: [noReplace] }).success).toBe(false)
      expect(response.safeParse({ status: 'done', proposals: [{ ...comment, comment: '' }] }).success).toBe(
        false,
      )
      expect(response.safeParse({ status: 'done', proposals: [{ ...edit, paragraph: 0 }] }).success).toBe(
        false,
      )
      expect(
        response.safeParse({ status: 'done', proposals: [{ ...edit, find: 'x'.repeat(2001) }] }).success,
      ).toBe(false)
      const many = Array.from({ length: MAX_AI_PROPOSALS + 1 }, (_, i) => ({ ...edit, id: String(i) }))
      expect(response.safeParse({ status: 'done', proposals: many }).success).toBe(false)
    })

    it('accepts an insertion at the start, and refuses one without content', () => {
      const insert = { id: '3', kind: 'insert', paragraph: 0, paragraphText: '', content: '# Título' }
      expect(response.safeParse({ status: 'done', proposals: [insert] }).success).toBe(true)
      expect(response.safeParse({ status: 'done', proposals: [{ ...insert, content: '' }] }).success).toBe(
        false,
      )
      expect(response.safeParse({ status: 'done', proposals: [{ ...insert, paragraph: -1 }] }).success).toBe(
        false,
      )
    })
  })

  it('accepts only an http or https address for the provider', () => {
    const save = ipcContracts[IpcChannel.AiProviderSave].request
    const draft = { kind: 'openai-compatible', label: 'Ollama' }
    expect(save.safeParse({ ...draft, baseUrl: 'http://localhost:11434/v1' }).success).toBe(true)
    expect(save.safeParse({ ...draft, baseUrl: 'file:///etc/passwd' }).success).toBe(false)
    expect(save.safeParse({ ...draft, baseUrl: 'javascript:alert(1)' }).success).toBe(false)
  })

  it('does not let the key return to the renderer in the provider list', () => {
    const list = ipcContracts[IpcChannel.AiProvidersList].response
    const parsed = list.parse({
      providers: [
        {
          id: 'p1',
          kind: 'anthropic',
          label: 'A',
          hasKey: true,
          local: false,
          consented: false,
          apiKey: 's',
        },
      ],
      defaultProviderId: 'p1',
      preferences: { keepFormatting: true, anthropicFallback: false, chatProposals: false, mcpServer: false },
    })
    expect(parsed.providers[0]).not.toHaveProperty('apiKey')
  })
})

describe('local MCP server contracts', () => {
  const requestId = '7d3f1c9e-8a2b-4c5d-9e0f-1a2b3c4d5e6f'

  it('the window answers the snapshot, the apply result or the refusal, and nothing else', () => {
    const reply = ipcContracts[IpcChannel.McpDocumentReply].request
    expect(reply.safeParse({ requestId, ask: 'snapshot', paragraphs: ['a'] }).success).toBe(true)
    expect(reply.safeParse({ requestId, ask: 'apply', applied: true }).success).toBe(true)
    expect(reply.safeParse({ requestId, ask: 'refused', refusal: 'read-only' }).success).toBe(true)
    expect(reply.safeParse({ requestId, ask: 'refused', refusal: 'outra' }).success).toBe(false)
    expect(reply.safeParse({ requestId, ask: 'save' }).success).toBe(false)
  })

  it('the client list has at most two short names', () => {
    const clients = pushContracts[IpcChannel.McpClients]
    expect(clients.safeParse({ clients: ['a', 'b'] }).success).toBe(true)
    expect(clients.safeParse({ clients: ['a', 'b', 'c'] }).success).toBe(false)
    expect(clients.safeParse({ clients: ['x'.repeat(65)] }).success).toBe(false)
  })
})
