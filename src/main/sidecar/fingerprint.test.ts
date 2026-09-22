/**
 * O contrato da impressão digital, com os dois lados de verdade.
 *
 * A gravação cirúrgica só funciona se o bloco que volta do editor for
 * **reconhecido** como o mesmo que o leitor produziu. Quem decide é
 * `Node.Fingerprint()`, no sidecar, comparando o que ele leu com o que o editor
 * devolveu — e entre uma coisa e outra o modelo atravessa o schema do
 * ProseMirror, que o reescreve: materializa todo atributo declarado, funde nós de
 * texto vizinhos com as mesmas marcas, ordena as marcas pela posição delas no
 * schema e descarta o que não conhece.
 *
 * Os testes de cada lado não pegam isso. `DocxRoundTripTests` clona o modelo pelo
 * JSON — fiel ao transporte, e cego ao schema; os testes do editor não têm
 * sidecar. Uma diferença que apareça só na costura faz a gravação regenerar o
 * documento inteiro **sem falhar em lugar nenhum**, que é o risco nº 1 do plano
 * técnico: o arquivo sai parecido e perde tudo o que não sabemos escrever.
 *
 * Aqui os dois lados se encontram: o documento é lido pelo sidecar publicado,
 * passa pelo schema montado com as extensões reais do editor, e volta ao sidecar
 * para ser gravado. São dois critérios, e o segundo é que dá sentido ao primeiro:
 * **zero blocos reescritos**, e as duas árvores **iguais** no subconjunto que a
 * impressão digital compara. Só o número deixaria passar a perda que o schema e a
 * impressão digital ignorassem juntos — foi comparando as árvores que se
 * descobriu que o editor devolve `colspan` e `rowspan` em toda célula, o que
 * fazia toda tabela ser regenerada ao salvar.
 *
 * Sem `Editor` do Tiptap, e não por preguiça: ele precisa de DOM, e não há
 * happy-dom nem jsdom instalados neste projeto — trazer um só para este teste
 * seria uma dependência a mais para vigiar. O que o `Editor` faz com um modelo
 * recebido em `content` é exatamente `Node.fromJSON` sobre o schema das
 * extensões (ver `createDocument` do Tiptap), e o que `getJSON()` devolve é o
 * `toJSON()` do documento — os dois passos que estão aqui.
 *
 * A lista de extensões é carregada por caminho em variável, e não por `import`
 * comum, por causa das fronteiras que o projeto leva a sério: `tsconfig.node.json`
 * não conhece `src/renderer` — o processo main não tem DOM —, e um `import`
 * estático o obrigaria a conhecer. O caminho em variável mantém a verificação de
 * tipos de cada camada como está e ainda usa, em tempo de execução, as extensões
 * **de verdade** do editor: uma lista copiada aqui envelheceria em silêncio, que
 * é exatamente o defeito que este teste existe para pegar.
 */

import { access, constants } from 'node:fs/promises'
import { promisify } from 'node:util'
import { inflateRaw } from 'node:zlib'
import { afterAll, describe, expect, it } from 'vitest'
import { getSchema, type Extensions } from '@tiptap/core'
import { Node as ProseMirrorNode } from '@tiptap/pm/model'
import {
  docxWithBulletList,
  docxWithComment,
  docxWithDescribedImage,
  docxWithHeaderGrid,
  docxWithSpacingOnBothSides,
  docxWithStretchedImage,
  docxWithStyledCells,
  docxWithTable,
  docxWithTextBox,
  docxWithVerticalAlignment,
  docxWithoutExtras,
} from '../../../e2e/fixtures.js'
import {
  DEFAULT_PARAGRAPH_DRAFT,
  LineSpacingKind,
  TextAlignment,
  paragraphAttrsFrom,
  paragraphDraftFrom,
} from '@services/document/paragraph-format.js'
import { SidecarClient } from './client.js'
import { SidecarMethod } from './protocol.js'
import { sidecarPathIn } from './locate.js'

const executable = sidecarPathIn(process.cwd())
const published = await access(executable, constants.X_OK).then(
  () => true,
  () => false,
)

const client = new SidecarClient(() => Promise.resolve(executable))
afterAll(() => client.dispose())

const editorExtensions = '../../renderer/document/editor-extensions.js'
const { buildEditorExtensions } = (await import(editorExtensions)) as {
  buildEditorExtensions: (onSearchStatusChange: (status: unknown) => void) => Extensions
}

const schema = getSchema(buildEditorExtensions(() => {}))

const inflate = promisify(inflateRaw)

/**
 * O `word/document.xml` de dentro do `.docx` gravado, sem descompactar em disco.
 *
 * Varre os cabeçalhos locais do ZIP, que é o mesmo caminho de `formatacao.spec.ts`
 * e por que motivo: trazer uma biblioteca de ZIP para ler um arquivo num teste
 * seria dependência a mais para auditar num aplicativo que se orgulha de ter
 * poucas.
 */
async function documentXmlOf(zip: Uint8Array): Promise<string> {
  const bytes = Buffer.from(zip)

  for (let i = 0; i + 30 <= bytes.length; i++) {
    if (bytes.readUInt32LE(i) !== 0x04034b50) continue

    const method = bytes.readUInt16LE(i + 8)
    const compressed = bytes.readUInt32LE(i + 18)
    const nameLength = bytes.readUInt16LE(i + 26)
    const extraLength = bytes.readUInt16LE(i + 28)
    const name = bytes.subarray(i + 30, i + 30 + nameLength).toString('utf8')
    if (name !== 'word/document.xml') continue

    const start = i + 30 + nameLength + extraLength
    const data = bytes.subarray(start, start + compressed)
    return method === 0 ? data.toString('utf8') : (await inflate(data)).toString('utf8')
  }

  throw new Error('o arquivo gravado não tem word/document.xml')
}

/** O modelo depois de atravessar o schema do editor, como `getJSON()` o devolve. */
function throughEditor(doc: unknown): unknown {
  return ProseMirrorNode.fromJSON(schema, doc).toJSON()
}

/**
 * O modelo reduzido ao que a impressão digital compara.
 *
 * É o espelho de `Node.Fingerprint()` no sidecar (`Nodes.cs`), e escrito à mão
 * de propósito: se os dois fossem o mesmo código, um erro no normalizador
 * cancelaria a si mesmo nos dois lados e o teste passaria em cima de uma perda.
 *
 * As quatro normalizações são as de lá, e cada uma é diferença **de forma**:
 * o `oid` é identidade e não conteúdo; atributo nulo é o mesmo que atributo
 * ausente, porque o ProseMirror materializa todo atributo do schema; a ordem das
 * marcas é do schema num lado e do `w:rPr` no outro; e nós de texto vizinhos com
 * as mesmas marcas são um só texto para o ProseMirror.
 *
 * O que ela **não** apaga é o que este teste existe para pegar: um atributo que
 * o leitor emite e o schema não conhece desaparece de um lado só, e a comparação
 * acusa. Comparar o modelo com ele mesmo depois de duas idas pelo schema não
 * acusava nada — a segunda ida não tem mais nada a perder.
 */
function asFingerprinted(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(asFingerprinted)
  if (value === null || typeof value !== 'object') return value

  const node: Record<string, unknown> = {}

  for (const [key, entry] of Object.entries(value as Record<string, unknown>)) {
    if (key === 'attrs') {
      const attrs = Object.entries((entry ?? {}) as Record<string, unknown>).filter(
        ([name, item]) => name !== 'oid' && item !== null && item !== undefined,
      )
      if (attrs.length > 0) node['attrs'] = Object.fromEntries(attrs.map(([name, item]) => [name, item]))
      continue
    }

    if (key === 'marks' && Array.isArray(entry)) {
      node['marks'] = entry
        .map(asFingerprinted)
        .sort((left, right) => typeOf(left).localeCompare(typeOf(right)))
      continue
    }

    if (key === 'content' && Array.isArray(entry)) {
      node['content'] = mergeNeighbouringText(entry.map(asFingerprinted))
      continue
    }

    node[key] = asFingerprinted(entry)
  }

  return node
}

function typeOf(node: unknown): string {
  const type = (node as { type?: unknown } | null)?.type
  return typeof type === 'string' ? type : ''
}

/** Nós de texto vizinhos com as mesmas marcas são um texto só. */
function mergeNeighbouringText(content: unknown[]): unknown[] {
  const merged: Array<Record<string, unknown>> = []

  for (const child of content) {
    const node = child as Record<string, unknown>
    const previous = merged.at(-1)

    if (
      previous !== undefined &&
      node['type'] === 'text' &&
      previous['type'] === 'text' &&
      stable(previous['marks']) === stable(node['marks'])
    ) {
      previous['text'] = `${String(previous['text'] ?? '')}${String(node['text'] ?? '')}`
      continue
    }

    merged.push({ ...node })
  }

  return merged
}

/** JSON com as chaves em ordem estável, para comparar marcas como texto. */
function stable(value: unknown): string {
  if (Array.isArray(value)) return `[${value.map(stable).join(',')}]`
  if (value === null || typeof value !== 'object') return JSON.stringify(value) ?? 'null'

  const entries = Object.entries(value as Record<string, unknown>).sort(([left], [right]) =>
    left.localeCompare(right),
  )
  return `{${entries.map(([key, entry]) => `${key}:${stable(entry)}`).join(',')}}`
}

interface OpenReply {
  readonly model: { readonly page: unknown; readonly doc: unknown }
}

interface SaveReply {
  readonly preservedBlocks: number
  readonly rewrittenBlocks: number
}

async function openSaveAndCount(bytes: Buffer): Promise<SaveReply> {
  const opened = await client.request(SidecarMethod.DocxOpen, {}, new Uint8Array(bytes))
  const { model } = opened.result as OpenReply

  const saved = await client.request(
    SidecarMethod.DocxSave,
    { page: model.page, doc: throughEditor(model.doc) },
    new Uint8Array(bytes),
  )

  return saved.result as SaveReply
}

describe.skipIf(!published)('impressão digital entre o editor e o sidecar', () => {
  // Um caso por estrutura do corpus, porque cada uma tem um jeito próprio de
  // divergir: o texto partido em runs, a imagem no fluxo, a caixa de texto
  // ancorada, o espaçamento sempre declarado, o cabeçalho em grade.
  const documents: Array<[string, () => Promise<Buffer>]> = [
    ['parágrafos com comentário ancorado', docxWithComment],
    ['parágrafo sem nada em volta', docxWithoutExtras],
    ['imagem esticada no fluxo do texto', docxWithStretchedImage],
    ['caixa de texto', docxWithTextBox],
    ['espaçamento dos dois lados', docxWithSpacingOnBothSides],
    ['cabeçalho em grade', docxWithHeaderGrid],
    ['lista com marcador', docxWithBulletList],
    ['tabela com tabela aninhada', docxWithTable],
    ['sobrescrito e subscrito', docxWithVerticalAlignment],
    // Os atributos do M4: largura de coluna, sombreamento, borda, mesclagem
    // horizontal e linha de cabeçalho na tabela; texto alternativo na imagem.
    ['tabela com sombreamento, borda e cabeçalho', docxWithStyledCells],
    ['imagem com texto alternativo', docxWithDescribedImage],
  ]

  it.each(documents)('abrir e salvar %s não reescreve bloco nenhum', async (_name, build) => {
    const bytes = await build()
    const result = await openSaveAndCount(bytes)

    expect(result.rewrittenBlocks).toBe(0)
    expect(result.preservedBlocks).toBeGreaterThan(0)
  })

  it.each(documents)('o modelo de %s volta do schema como o sidecar o leu', async (_name, build) => {
    // O número acima diz que a comparação passou; este diz **por que** —
    // comparando as duas árvores pelo mesmo subconjunto que a impressão digital
    // usa. Sem isto, uma perda que o schema e a impressão digital ignorassem
    // igualmente passava pelos dois: o contador daria zero e nada seria pego.
    const bytes = await build()
    const opened = await client.request(SidecarMethod.DocxOpen, {}, new Uint8Array(bytes))
    const { model } = opened.result as OpenReply

    expect(asFingerprinted(throughEditor(model.doc))).toEqual(asFingerprinted(model.doc))
  })

  it('a saída do diálogo de parágrafo é o que o arquivo recebe', async () => {
    // A gravação passou a receber esta forma: `lineHeight` como medida de CSS,
    // `textAlign` e o nível de recuo zerado (`indent: 0`). É o único caminho em que
    // o editor **inventa** atributos de parágrafo em vez de devolver os que leu, e
    // é onde o erro de entrelinha morava: o número do diálogo ia cru para o
    // atributo, o gravador dividia pela altura natural da fonte e o arquivo
    // recebia 1,23 linha onde a pessoa pediu 1,5.
    const bytes = await docxWithoutExtras()
    const opened = await client.request(SidecarMethod.DocxOpen, {}, new Uint8Array(bytes))
    const { model } = opened.result as OpenReply

    const doc = model.doc as { content: Array<{ type: string; attrs: Record<string, unknown> }> }
    const first = doc.content[0]!
    const attrs = paragraphAttrsFrom(
      {
        ...DEFAULT_PARAGRAPH_DRAFT,
        align: TextAlignment.Justify,
        lineSpacingKind: LineSpacingKind.Multiple,
        lineSpacingValue: 1.5,
        spaceAfter: 6,
      },
      first.attrs,
    )
    const edited = {
      ...(model.doc as Record<string, unknown>),
      content: [{ ...first, attrs: { ...first.attrs, ...attrs } }, ...doc.content.slice(1)],
    }

    const saved = await client.request(
      SidecarMethod.DocxSave,
      { page: model.page, doc: throughEditor(edited) },
      new Uint8Array(bytes),
    )
    const result = saved.result as SaveReply & { inventory: { lost: string[] } }

    // Um bloco reescrito — o que passou pelo diálogo — e nenhuma perda: entrelinha
    // fora da faixa que o gravador aceita viraria linha no inventário e nada no
    // arquivo, sem ninguém avisar.
    expect(result.rewrittenBlocks).toBe(1)
    expect(result.inventory.lost).toEqual([])

    // O número no arquivo, que é o que o Word lê: `w:line` em 240-avos, com
    // `w:lineRule="auto"`. 1,5 linha são 360. Com o diálogo escrevendo o fator cru
    // no atributo, o gravador dividia 1,5 pela altura natural da fonte e gravava
    // 313 — 1,3 linha, e ninguém era avisado.
    const xml = await documentXmlOf(saved.binary)
    expect(xml).toContain('w:line="360"')

    // E a prova da ida e volta: reabrir devolve ao diálogo o que ele escolheu.
    const reopened = await client.request(SidecarMethod.DocxOpen, {}, saved.binary)
    const back = (reopened.result as OpenReply).model.doc as {
      content: Array<{ attrs: Record<string, unknown> }>
    }
    const draft = paragraphDraftFrom(back.content[0]!.attrs)

    expect(draft.lineSpacingKind).toBe(LineSpacingKind.Multiple)
    expect(draft.lineSpacingValue).toBe(1.5)
    expect(draft.align).toBe(TextAlignment.Justify)
    expect(draft.spaceAfter).toBe(6)
  })

  it('cada bloco do modelo chega ao editor com identidade', async () => {
    // O `oid` é o que liga o bloco da tela ao `w:p` do arquivo. Sem ele não há
    // gravação cirúrgica nenhuma, e a comparação acima não teria o que comparar.
    const bytes = await docxWithComment()
    const opened = await client.request(SidecarMethod.DocxOpen, {}, new Uint8Array(bytes))
    const { model } = opened.result as OpenReply

    const blocks = (model.doc as { content: Array<{ attrs?: Record<string, unknown> }> }).content
    expect(blocks.length).toBeGreaterThan(0)
    expect(blocks.every((block) => typeof block.attrs?.['oid'] === 'string')).toBe(true)
  })
})
