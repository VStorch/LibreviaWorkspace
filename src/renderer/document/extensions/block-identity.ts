import { Extension } from '@tiptap/core'
import { Plugin, PluginKey, type Transaction } from '@tiptap/pm/state'

/**
 * O que o bloco traz do arquivo e o editor não interpreta — identidade e
 * objetos ancorados.
 *
 * O leitor do sidecar carimba um `oid` em cada bloco de primeiro nível
 * (`BodyReader.NewBlock`), e a gravação usa esse `oid` para decidir o que
 * **não** reescrever: bloco cujo conteúdo não mudou volta para o `.docx` como o
 * XML original, byte a byte (`DocxWriter.OidOf`). É o eixo inteiro da gravação
 * cirúrgica, e o que protege comentário, revisão, caixa de texto e forma de
 * sumirem num documento que o editor não sabe reproduzir.
 *
 * O ProseMirror descarta todo atributo que o schema não declara. Sem esta
 * extensão o `oid` morria na travessia pelo editor, `OidOf` devolvia `null` para
 * cada bloco e a gravação degradava para regeneração completa — **em silêncio**,
 * que é o modo de falha que o plano técnico chama de risco nº 1 (§6.1).
 *
 * Medido antes da correção, abrindo e salvando `modelo-de-manual.docx`
 * **sem editar nada**: `document.xml` caiu de 32.282 para 5.094 bytes, os 23
 * `w14:paraId` viraram 0 e as quatro caixas de texto (16 `txbxContent`)
 * desapareceram do arquivo. Não era perda de formatação: era perda de conteúdo,
 * num arquivo que o usuário só abriu para ler.
 *
 * Vai para o HTML como `data-oid` — e não só para o JSON — porque a identidade
 * também precisa atravessar recortar/colar e desfazer, que passam pelo DOM. E
 * porque o DOM a duplica, a identidade repetida é desfeita aqui mesmo, por
 * `uniqueOids`.
 */

export interface BlockIdentityOptions {
  types: string[]
}

/**
 * Um `oid` por bloco: a segunda ocorrência perde a identidade.
 *
 * Dois blocos com o mesmo `oid` é o pior caso da gravação cirúrgica. O gravador
 * preserva o XML original na **primeira** ocorrência e regenera as demais — de
 * modo que o bloco que a pessoa nem tocou volta reescrito, e com ele se vai o que
 * o editor não sabe reproduzir. Foi assim que o marcador de um parágrafo dividido
 * desapareceu do arquivo.
 *
 * Quem duplica é a divisão de parágrafo: o Enter entrega os dois lados com os
 * atributos do original, `oid` incluído. A colagem faz o mesmo, pelo `data-oid`
 * do HTML. A regra aqui é a que o gravador já aplica, um passo antes e uma vez
 * só: o primeiro fica com a identidade — é ele que está no lugar do bloco que
 * veio do arquivo — e o novo nasce sem nenhuma, o que o gravador entende como
 * "bloco novo, gere do zero".
 *
 * Só quando há o que corrigir, e sem descer dentro do parágrafo: uma transação
 * apendada a cada tecla suja o histórico de desfazer, e o percurso das palavras
 * do documento não paga nada a esta conta.
 */
export function uniqueOids(): Plugin {
  return new Plugin({
    key: new PluginKey('blockIdentityUnique'),

    appendTransaction(transactions, _oldState, newState) {
      if (!transactions.some((transaction) => transaction.docChanged)) return null

      const seen = new Set<string>()
      let corrections: Transaction | null = null

      newState.doc.descendants((node, position) => {
        const oid: unknown = node.attrs['oid']
        if (typeof oid === 'string' && oid.length > 0) {
          if (seen.has(oid)) {
            corrections ??= newState.tr
            corrections.setNodeAttribute(position, 'oid', null)
          } else {
            seen.add(oid)
          }
        }

        // Dentro de um parágrafo não há bloco com identidade: o que mora lá é
        // texto, e percorrê-lo custaria o documento inteiro a cada tecla.
        return !node.isTextblock
      })

      return corrections
    },
  })
}

export const BlockIdentity = Extension.create<BlockIdentityOptions>({
  name: 'blockIdentity',

  addProseMirrorPlugins() {
    return [uniqueOids()]
  },

  addOptions() {
    // Exatamente os nós em que `BodyReader` chama `NewBlock`: o parágrafo de
    // topo (que pode sair como `paragraph`, `heading` ou `pageBreak`), o item de
    // lista — que no arquivo é um `w:p` — e a tabela. Declarar em mais nós não
    // machucaria, mas sugeriria uma identidade que o leitor não emite.
    return { types: ['paragraph', 'heading', 'pageBreak', 'listItem', 'table'] }
  },

  addGlobalAttributes() {
    return [
      {
        types: this.options.types,
        attributes: {
          oid: {
            default: null,
            parseHTML: (element) => element.getAttribute('data-oid'),
            renderHTML: (attributes) => {
              const oid = attributes['oid']
              return typeof oid === 'string' && oid.length > 0 ? { 'data-oid': oid } : {}
            },
          },

          /**
           * Objetos ancorados neste bloco: imagem ou caixa de texto que não
           * estão no fluxo.
           *
           * Dado opaco, como o `oid` — nada aqui os interpreta, e é por isso que
           * precisam ser declarados: o ProseMirror descarta atributo fora do
           * schema, e a capa perderia a marca e as caixas ao atravessar o
           * editor, inclusive ao salvar.
           *
           * Declarado nos mesmos nós que o `oid`, e não só em parágrafo e
           * título: quando um parágrafo tem só a imagem e uma quebra de página,
           * ele vira o nó `pageBreak` — e os objetos dele iam junto.
           *
           * Não vai para o HTML: é dado, não aparência. Quem desenha lê do
           * modelo e põe numa camada própria, fora do texto editável.
           */
          floats: {
            default: null,
            parseHTML: () => null,
            renderHTML: () => ({}),
          },
        },
      },
    ]
  },
})
