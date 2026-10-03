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

/** O CSS dela está em `content-styles.ts`. */
export const MATH_CLASS = 'equacao'

/**
 * `m:oMath` ou `m:oMathPara`, atômico como o campo. A identidade é o `omml`, que
 * volta ao arquivo; o `mathml`, a lista `lossy` e o `editable` saem dele no
 * sidecar (`OmmlMath.cs`). O `latex` é a fonte da edição: a equação editada chega
 * com `omml` nulo, e o sidecar o refaz do MathML. A de exibição fica no
 * parágrafo, como no OOXML. O MathML passa sempre por `sanitizeMathMl` e vira DOM
 * por `createElementNS`, nunca por `innerHTML`.
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
      // O `data-omml` vazio é o da equação que não veio de arquivo: volta a ser nulo.
      omml: { default: null, parseHTML: (element) => element.getAttribute('data-omml') || null },
      // Nulo fica o do `getAttrs` da regra, do `math` colado de fora.
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

  // O embrulho é o que o editor copia; o `math` solto vem de fora e vira equação nova.
  parseHTML() {
    return [{ tag: 'span[data-math]' }, { tag: 'math', getAttrs: pastedMathAttrs }]
  },

  // Um elemento pronto: o MathML filtrado é montado nó a nó, para a tela, a impressão e a cópia.
  renderHTML({ node }) {
    return renderMath(node, document)
  },

  renderText({ node }) {
    return plainTextOf(node)
  },
})

/** O LaTeX, guardado ou tirado do MathML, ou o marcador quando nem isso dá. */
export function plainTextOf(node: ProseMirrorNode): string {
  const latex = latexOfEquation(node.attrs)
  return latex === '' ? t('document.math.placeholder') : latex
}

/** `false` quando o MathML não passa no filtro. */
export function pastedMathAttrs(element: HTMLElement): Record<string, unknown> | false {
  return mathAttrsOfMarkup(new XMLSerializer().serializeToString(element))
}

/** Sem OMML, com o LaTeX tirado do MathML. */
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

/** Também a visualização do editor de equações. */
export function buildMath(element: MathElement, doc: Document): Element {
  const built = doc.createElementNS(MATHML_NAMESPACE, element.tag)
  for (const [name, value] of Object.entries(element.attrs)) built.setAttribute(name, value)
  for (const child of element.children) {
    built.append(typeof child === 'string' ? doc.createTextNode(child) : buildMath(child, doc))
  }
  return built
}

/**
 * O clique duplo e o Enter chegam a `EditEquation`, que decide se abre para
 * editar ou só para ver. Prioridade alta, para o Enter vir antes do que parte o
 * parágrafo, sem mexer na ordem do esquema.
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
            // A posição do nó: na metade direita, o clique cai depois da equação.
            view.dispatch(view.state.tr.setSelection(NodeSelection.create(view.state.doc, nodePos)))
            emitEditorCommand(EditorCommand.EditEquation)
            return true
          },
        },
      }),
    ]
  },
})
