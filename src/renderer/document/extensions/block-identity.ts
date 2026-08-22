import { Extension } from '@tiptap/core'

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
 * também precisa atravessar recortar/colar e desfazer, que passam pelo DOM. Um
 * `oid` repetido por colagem é previsto do outro lado: a gravação preserva o XML
 * original na primeira ocorrência e regenera as demais.
 */

export interface BlockIdentityOptions {
  types: string[]
}

export const BlockIdentity = Extension.create<BlockIdentityOptions>({
  name: 'blockIdentity',

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
