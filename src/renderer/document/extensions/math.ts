import { Extension, Node } from '@tiptap/core'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import { NodeSelection, Plugin, PluginKey } from '@tiptap/pm/state'
import { EditorCommand, emitEditorCommand } from '../editor-commands.js'
import { equationAtSelection } from '../math-commands.js'
import {
  MATHML_NAMESPACE,
  mathMlToString,
  sanitizeMathMl,
  type MathElement,
} from '@services/document/mathml.js'
import { latexOfEquation } from '@services/document/mathml-latex.js'
import { t } from '../../i18n.js'

/** A classe do embrulho — o CSS dela está em `content-styles.ts`. */
export const MATH_CLASS = 'equacao'

/**
 * Uma equação do Word (`m:oMath`, ou `m:oMathPara` na de exibição) — M11, fase 1.
 *
 * Um nó atômico, como o campo: o cursor passa por ele de uma vez, o Backspace o
 * apaga inteiro e arrastar o leva junto. A identidade é o `omml`, o XML que o
 * arquivo trazia e que volta a ele na gravação; o resto sai dele no sidecar
 * (`OmmlMath.cs`): o `mathml` que a tela desenha, a lista `lossy` do que ela não
 * desenha e o `editable`, falso quando a lista não é vazia. O `latex` é a fonte da
 * edição (fase 2, `MathDialog.tsx`): a equação nova ou editada chega com `omml`
 * nulo, e o sidecar refaz o OMML a partir do MathML.
 *
 * A de exibição continua no parágrafo dela, porque é ali que o OOXML a guarda; é
 * o desenho que a põe num bloco, alinhado pelo `jc`. Para a paginação ela é uma
 * linha só, alta, que nunca se parte — ver `line-boxes.ts`.
 *
 * O MathML passa sempre por `sanitizeMathMl` e vira DOM por `createElementNS`:
 * nunca por `innerHTML`, porque o atributo mora num `.sdoc` que se edita à mão.
 */
export const MathNode = Node.create({
  name: 'math',
  group: 'inline',
  inline: true,
  atom: true,
  selectable: true,
  draggable: true,

  addAttributes() {
    return {
      // O `data-omml` vazio é o da equação que não veio de arquivo (o embrulho
      // escreve o nulo como texto): volta a ser nulo, e o sidecar refaz o OMML.
      omml: { default: null, parseHTML: (element) => element.getAttribute('data-omml') || null },
      // Nulo é "não diz": fica o do `getAttrs` da regra — o do `math` colado de
      // fora, que não tem os `data-*` — ou o padrão.
      mathml: { default: '', parseHTML: (element) => element.getAttribute('data-mathml') },
      latex: { default: '', parseHTML: (element) => element.getAttribute('data-latex') },
      display: {
        default: false,
        parseHTML: (element) =>
          element.hasAttribute('data-display') ? element.getAttribute('data-display') === 'true' : null,
      },
      jc: { default: null, parseHTML: (element) => element.getAttribute('data-jc') },
      lossy: { default: [], parseHTML: (element) => lossyOf(element.getAttribute('data-lossy')) },
      editable: { default: true, parseHTML: (element) => element.getAttribute('data-editable') !== 'false' },
    }
  },

  // O embrulho é o que o próprio editor copia, com o `omml` junto. O `math` solto
  // é o que vem de fora — uma página da Web, a nossa exportação em HTML —: vira
  // uma equação nova, só com o MathML filtrado e o LaTeX tirado dele.
  parseHTML() {
    return [{ tag: 'span[data-math]' }, { tag: 'math', getAttrs: pastedMathAttrs }]
  },

  // Um elemento pronto, e não uma especificação: o MathML filtrado é montado nó a
  // nó. Serve à tela, à impressão (que serializa pelos nós — `print-source.ts`) e
  // à área de transferência, que leva os atributos para colar a equação inteira.
  renderHTML({ node }) {
    return renderMath(node, document)
  },

  renderText({ node }) {
    return plainTextOf(node)
  },
})

/**
 * O texto que a equação deixa ao ser copiada: o LaTeX, o guardado ou o tirado do
 * MathML, ou o marcador quando nem isso dá.
 */
export function plainTextOf(node: ProseMirrorNode): string {
  const latex = latexOfEquation(node.attrs)
  return latex === '' ? t('document.math.placeholder') : latex
}

/** Os atributos do `math` colado de fora, ou `false` quando o MathML não passa no filtro. */
export function pastedMathAttrs(element: HTMLElement): Record<string, unknown> | false {
  return mathAttrsOfMarkup(new XMLSerializer().serializeToString(element))
}

/** A equação nova que sai de um MathML de fora: sem OMML, com o LaTeX tirado dele. */
export function mathAttrsOfMarkup(markup: string): Record<string, unknown> | false {
  const tree = sanitizeMathMl(markup)
  if (tree === null) return false
  const mathml = mathMlToString(tree)
  return {
    omml: null,
    mathml,
    latex: latexOfEquation({ mathml }),
    display: tree.attrs['display'] === 'block',
  }
}

function lossyOf(raw: string | null): string[] {
  if (raw === null || raw === '') return []
  try {
    const parsed: unknown = JSON.parse(raw)
    return Array.isArray(parsed) ? parsed.filter((item): item is string => typeof item === 'string') : []
  } catch {
    return []
  }
}

/** O embrulho da equação, com o MathML filtrado dentro. */
export function renderMath(node: ProseMirrorNode, doc: Document): HTMLElement {
  const display = node.attrs['display'] === true
  const lossy = (node.attrs['lossy'] as readonly string[] | null) ?? []
  const locked = node.attrs['editable'] === false || lossy.length > 0
  const jc = typeof node.attrs['jc'] === 'string' ? node.attrs['jc'] : null

  const wrapper = doc.createElement('span')
  wrapper.className = [
    MATH_CLASS,
    display ? `${MATH_CLASS}--exibicao` : '',
    locked ? `${MATH_CLASS}--travada` : '',
  ]
    .filter((name) => name !== '')
    .join(' ')
  wrapper.setAttribute('data-math', '')
  wrapper.setAttribute('data-omml', String(node.attrs['omml'] ?? ''))
  wrapper.setAttribute('data-mathml', String(node.attrs['mathml'] ?? ''))
  wrapper.setAttribute('data-latex', String(node.attrs['latex'] ?? ''))
  wrapper.setAttribute('data-display', String(display))
  if (jc !== null) wrapper.setAttribute('data-jc', jc)
  wrapper.setAttribute('data-lossy', JSON.stringify(lossy))
  wrapper.setAttribute('data-editable', String(!locked))
  wrapper.title = locked
    ? t('document.math.locked', { constructs: lossy.join(', ') })
    : t('document.math.title')

  const tree = sanitizeMathMl(String(node.attrs['mathml'] ?? ''))
  if (tree === null) wrapper.textContent = plainTextOf(node)
  else wrapper.append(buildMath(tree, doc))
  return wrapper
}

/** O MathML já filtrado em DOM, nó a nó — também a visualização do editor de equações. */
export function buildMath(element: MathElement, doc: Document): Element {
  const built = doc.createElementNS(MATHML_NAMESPACE, element.tag)
  for (const [name, value] of Object.entries(element.attrs)) built.setAttribute(name, value)
  for (const child of element.children) {
    built.append(typeof child === 'string' ? doc.createTextNode(child) : buildMath(child, doc))
  }
  return built
}

/**
 * Abrir a equação no editor: o clique duplo nela, ou o Enter com ela selecionada
 * (M11, fase 2). Os dois chegam ao mesmo comando do menu, `EditEquation`, que é
 * quem decide se abre para editar ou só para ver.
 *
 * Extensão à parte, com prioridade alta, para o Enter dela vir antes do Enter
 * que parte o parágrafo — e sem mexer na prioridade do nó, que decide a ordem
 * do esquema.
 */
export const MathEditing = Extension.create({
  name: 'mathEditing',
  priority: 1000,

  addKeyboardShortcuts() {
    return {
      Enter: () => {
        if (equationAtSelection(this.editor.state) === null) return false
        emitEditorCommand(EditorCommand.EditEquation)
        return true
      },
    }
  },

  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: new PluginKey('mathEditing'),
        props: {
          // A assinatura é a do ProseMirror, e não nossa.
          // eslint-disable-next-line max-params
          handleDoubleClickOn(view, _pos, node, nodePos, _event, direct) {
            if (!direct || node.type.name !== 'math') return false
            // A posição do nó, e não a do clique: na metade direita da equação o
            // clique cai depois dela, onde não há nó para selecionar.
            view.dispatch(view.state.tr.setSelection(NodeSelection.create(view.state.doc, nodePos)))
            emitEditorCommand(EditorCommand.EditEquation)
            return true
          },
        },
      }),
    ]
  },
})
