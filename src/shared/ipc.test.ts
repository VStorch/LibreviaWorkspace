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

describe('contratos de IPC', () => {
  it('define um contrato para cada canal invocável', () => {
    // A channel without a schema would pass unvalidated payloads to the handler.
    for (const channel of INVOCABLE_IPC_CHANNELS) {
      expect(ipcContracts[channel]).toBeDefined()
    }
  })

  it('não expõe contrato para canal não declarado', () => {
    expect(Object.keys(ipcContracts).sort()).toEqual([...INVOCABLE_IPC_CHANNELS].sort())
  })

  it('mantém o canal de menu fora dos invocáveis: ele vai de main para renderer', () => {
    expect(INVOCABLE_IPC_CHANNELS).not.toContain(IpcChannel.MenuCommand)
    expect(ipcContracts).not.toHaveProperty(IpcChannel.MenuCommand)
  })
})

describe('contratos do sentido main → renderer', () => {
  it('define um contrato para cada canal empurrado', () => {
    // Without a schema nobody would check the shape, and the renderer would drop the message
    // silently, the costliest defect here.
    for (const channel of PUSH_IPC_CHANNELS) {
      expect(pushContracts[channel]).toBeDefined()
    }
  })

  it('os dois sentidos não se misturam', () => {
    for (const channel of PUSH_IPC_CHANNELS) {
      expect(INVOCABLE_IPC_CHANNELS).not.toContain(channel)
      expect(ipcContracts).not.toHaveProperty(channel)
    }
  })
})

describe('validação do alvo do menu de contexto', () => {
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

  it('aceita o que o Chromium manda', () => {
    expect(schema.safeParse(alvo).success).toBe(true)
  })

  it('recusa coordenada negativa e lista de sugestões absurda', () => {
    // This becomes a screen position and menu items: a negative coordinate puts the menu outside
    // the window, and thirty suggestions overflow it.
    expect(schema.safeParse({ ...alvo, x: -1 }).success).toBe(false)
    expect(
      schema.safeParse({ ...alvo, dictionarySuggestions: Array.from({ length: 30 }, () => 'x') }).success,
    ).toBe(false)
  })
})

describe('validação de prefs:set', () => {
  const schema = ipcContracts[IpcChannel.PreferencesSet].request

  it('aceita um remendo de uma chave só', () => {
    // Whoever toggles formatting marks has no opinion about spelling.
    expect(schema.parse({ invisibleCharacters: true })).toEqual({ invisibleCharacters: true })
  })

  it('recusa valor que não é booleano', () => {
    expect(schema.safeParse({ spellcheck: 'sim' }).success).toBe(false)
  })
})

describe('validação de file:save', () => {
  const schema = ipcContracts[IpcChannel.FileSave].request

  it('aceita uma gravação bem formada', () => {
    expect(schema.safeParse({ path: '/home/ana/ata.txt', content: 'texto', origin: null }).success).toBe(true)
  })

  it('aceita a origem do documento aberto', () => {
    const payload = { path: '/home/ana/ata.docx', content: '{}', origin: '/home/ana/ata.sdoc' }
    expect(schema.safeParse(payload).success).toBe(true)
  })

  it.each([
    ['sem caminho', { content: 'texto' }],
    ['caminho vazio', { path: '', content: 'texto' }],
    ['sem conteúdo', { path: '/a/b.txt' }],
    ['caminho não textual', { path: 42, content: 'texto', origin: null }],
    ['sem origem', { path: '/a/b.docx', content: 'texto' }],
    ['origem vazia', { path: '/a/b.docx', content: 'texto', origin: '' }],
    ['nulo', null],
  ])('recusa %s', (_label, payload) => {
    expect(schema.safeParse(payload).success).toBe(false)
  })

  it('recusa conteúdo acima do teto de memória', () => {
    const oversized = { path: '/a/b.txt', content: 'x'.repeat(MAX_TEXT_LENGTH + 1) }
    expect(schema.safeParse(oversized).success).toBe(false)
  })

  it('descarta campos não previstos no contrato', () => {
    const parsed = schema.parse({ path: '/a/b.txt', content: 'oi', origin: null, extra: 'ignorar' })
    expect(parsed).toEqual({ path: '/a/b.txt', content: 'oi', origin: null })
  })
})

describe('modelos do Word (M11)', () => {
  const open = ipcContracts[IpcChannel.TemplateOpen].request
  const opened = ipcContracts[IpcChannel.TemplateOpen].response
  const list = ipcContracts[IpcChannel.TemplateList].response

  it('abre pelo par fonte e id', () => {
    expect(open.safeParse({ source: 'builtin', id: 'carta.dotx' }).success).toBe(true)
    expect(open.safeParse({ source: 'user', id: '/home/ana/Modelos/proposta.dotx' }).success).toBe(true)
  })

  it.each([
    ['fonte desconhecida', { source: 'web', id: 'carta.dotx' }],
    ['id vazio', { source: 'builtin', id: '' }],
    ['sem id', { source: 'user' }],
  ])('recusa %s', (_label, payload) => {
    expect(open.safeParse(payload).success).toBe(false)
  })

  it('o documento aberto de um modelo traz a marca de modelo', () => {
    const file = {
      path: '/r/carta.dotx',
      name: 'Carta.docx',
      kind: 'document',
      content: '{}',
      template: true,
    }
    expect(opened.parse({ file }).file.template).toBe(true)
  })

  it('a galeria lista as duas fontes e a pasta do usuário', () => {
    const entry = { source: 'builtin', id: 'carta.dotx', name: 'Carta', description: 'Uma carta' }
    expect(list.safeParse({ builtin: [entry], user: [], folder: '/home/ana/Modelos' }).success).toBe(true)
    expect(list.safeParse({ builtin: [{ ...entry, source: 'outra' }], user: [], folder: '' }).success).toBe(
      false,
    )
  })
})

describe('validação de file:open-recent', () => {
  const schema = ipcContracts[IpcChannel.FileOpenRecent].request

  it('exige um caminho não vazio', () => {
    expect(schema.safeParse({ path: '' }).success).toBe(false)
    expect(schema.safeParse({ path: '/a/b.txt' }).success).toBe(true)
  })
})

describe('validação de fonts:list', () => {
  // `src/main/ipc/registry.ts` runs this schema on the real response, and `registry.test.ts` proves
  // it: a schema nobody runs is a compile-time type, not protection.
  const schema = ipcContracts[IpcChannel.FontsList].response

  it('aceita a lista vazia', () => {
    // A system without `fontconfig` returns nothing, and that is not an error: the toolbar keeps
    // the bundled fonts.
    expect(schema.safeParse({ families: [] }).success).toBe(true)
  })

  it('recusa nome vazio e lista absurda', () => {
    // This is the output of a system program, data and not truth: an empty name would become an
    // invisible option, and a list of a hundred thousand entries would freeze the toolbar.
    expect(schema.safeParse({ families: [''] }).success).toBe(false)
    expect(schema.safeParse({ families: Array.from({ length: 4001 }, () => 'Arial') }).success).toBe(false)
  })
})

describe('validação de dialog:confirm-discard', () => {
  const response = ipcContracts[IpcChannel.DialogConfirmDiscard].response

  it.each(['save', 'discard', 'cancel'])('aceita a resposta %s', (choice) => {
    expect(response.safeParse({ choice }).success).toBe(true)
  })

  it('recusa uma resposta fora das três previstas', () => {
    expect(response.safeParse({ choice: 'talvez' }).success).toBe(false)
  })
})

describe('contratos do assistente de IA', () => {
  const generate = ipcContracts[IpcChannel.AiGenerate].request
  const valid = {
    requestId: '0b8f3c1e-5d2a-4c7b-9e6f-1a2b3c4d5e6f',
    providerId: 'p1',
    model: 'llama3',
    action: 'rewrite',
    text: 'Um parágrafo.',
  }

  it('aceita um pedido completo', () => {
    expect(generate.safeParse(valid).success).toBe(true)
  })

  it('recusa seleção vazia, longa demais ou sem identificador válido', () => {
    expect(generate.safeParse({ ...valid, text: '' }).success).toBe(false)
    expect(generate.safeParse({ ...valid, text: 'x'.repeat(20_001) }).success).toBe(false)
    expect(generate.safeParse({ ...valid, requestId: 'abc' }).success).toBe(false)
    expect(generate.safeParse({ ...valid, action: 'delete-everything' }).success).toBe(false)
  })

  describe('planilha', () => {
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

    it('aceita cada tarefa com o que ela precisa', () => {
      expect(sheet.safeParse(formula).success).toBe(true)
      expect(sheet.safeParse({ ...formula, task: 'explain', instruction: '', formula: '=A1' }).success).toBe(
        true,
      )
      expect(sheet.safeParse({ ...formula, task: 'fill', rowCount: 500 }).success).toBe(true)
    })

    it('recusa tarefa sem o que precisa, linhas demais e contexto longo demais', () => {
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

  describe('conversa', () => {
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

    it('aceita uma conversa que termina com a pessoa', () => {
      expect(chat.safeParse(conversation).success).toBe(true)
      expect(chat.parse(conversation).proposals).toBe(false)
      expect(chat.safeParse({ ...conversation, proposals: 'sim' }).success).toBe(false)
      expect(chat.safeParse({ ...conversation, before: 'Doc', selection: 'trecho' }).success).toBe(true)
    })

    it('recusa papel inválido, conversa vazia ou que termina com o assistente', () => {
      expect(chat.safeParse({ ...conversation, messages: [{ role: 'system', text: 'x' }] }).success).toBe(
        false,
      )
      expect(chat.safeParse({ ...conversation, messages: [] }).success).toBe(false)
      expect(chat.safeParse({ ...conversation, messages: conversation.messages.slice(0, 2) }).success).toBe(
        false,
      )
    })

    it('recusa mensagens demais ou texto longo demais', () => {
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

    it('aceita os parágrafos numerados só dentro do tamanho do documento', () => {
      expect(chat.parse(conversation).paragraphs).toEqual([])
      expect(chat.safeParse({ ...conversation, paragraphs: ['Um.', 'Dois.'] }).success).toBe(true)
      const half = 'x'.repeat(MAX_AI_DOCUMENT_LENGTH / 2 + 1)
      expect(chat.safeParse({ ...conversation, paragraphs: [half, half] }).success).toBe(false)
    })
  })

  describe('propostas na resposta da conversa', () => {
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

    it('aceita alteração e comentário', () => {
      expect(response.safeParse({ status: 'done', proposals: [edit, comment] }).success).toBe(true)
      expect(response.safeParse({ status: 'done' }).success).toBe(true)
    })

    it('recusa proposta sem o texto que lhe cabe, parágrafo zero, trecho longo ou propostas demais', () => {
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
  })

  it('só aceita endereço http ou https para o provedor', () => {
    const save = ipcContracts[IpcChannel.AiProviderSave].request
    const draft = { kind: 'openai-compatible', label: 'Ollama' }
    expect(save.safeParse({ ...draft, baseUrl: 'http://localhost:11434/v1' }).success).toBe(true)
    expect(save.safeParse({ ...draft, baseUrl: 'file:///etc/passwd' }).success).toBe(false)
    expect(save.safeParse({ ...draft, baseUrl: 'javascript:alert(1)' }).success).toBe(false)
  })

  it('não deixa a chave voltar ao renderer na lista de provedores', () => {
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

describe('contratos do servidor MCP local', () => {
  const requestId = '7d3f1c9e-8a2b-4c5d-9e0f-1a2b3c4d5e6f'

  it('a janela responde o instantâneo, o resultado de aplicar ou a recusa, e nada além', () => {
    const reply = ipcContracts[IpcChannel.McpDocumentReply].request
    expect(reply.safeParse({ requestId, ask: 'snapshot', paragraphs: ['a'] }).success).toBe(true)
    expect(reply.safeParse({ requestId, ask: 'apply', applied: true }).success).toBe(true)
    expect(reply.safeParse({ requestId, ask: 'refused', refusal: 'read-only' }).success).toBe(true)
    expect(reply.safeParse({ requestId, ask: 'refused', refusal: 'outra' }).success).toBe(false)
    expect(reply.safeParse({ requestId, ask: 'save' }).success).toBe(false)
  })

  it('a lista de clientes tem no máximo dois nomes curtos', () => {
    const clients = pushContracts[IpcChannel.McpClients]
    expect(clients.safeParse({ clients: ['a', 'b'] }).success).toBe(true)
    expect(clients.safeParse({ clients: ['a', 'b', 'c'] }).success).toBe(false)
    expect(clients.safeParse({ clients: ['x'.repeat(65)] }).success).toBe(false)
  })
})
