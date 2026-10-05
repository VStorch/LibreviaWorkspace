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

/** Its CSS is in `content-styles.ts`. */
export const MATH_CLASS = 'equacao'

/**
 * `m:oMath` or `m:oMathPara`, atomic like a field. The identity is `omml`, which goes back to the
 * file; `mathml`, the `lossy` list and `editable` come from it in the sidecar (`OmmlMath.cs`).
 * `latex` is the editing source: an edited equation arrives with a null `omml`, and the sidecar
 * rebuilds it from the MathML. A display equation stays in the paragraph, as in OOXML. MathML
 * always goes through `sanitizeMathMl` and becomes DOM via `createElementNS`, never `innerHTML`.
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
      // An empty `data-omml` belongs to an equation that did not come from a file: it goes back to
      // null.
      omml: { default: null, parseHTML: (element) => element.getAttribute('data-omml') || null },
      // Null comes from the rule's `getAttrs`, for `math` pasted from outside.
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

  // The wrapper is what the editor copies; a loose `math` comes from outside and becomes a new
  // equation.
  parseHTML() {
    return [{ tag: 'span[data-math]' }, { tag: 'math', getAttrs: pastedMathAttrs }]
  },

  // A ready element: the filtered MathML is built node by node, for screen, print and copy.
  renderHTML({ node }) {
    return renderMath(node, document)
  },

  renderText({ node }) {
    return plainTextOf(node)
  },
})

/** The stored LaTeX or the one derived from MathML, or the placeholder when neither works. */
export function plainTextOf(node: ProseMirrorNode): string {
  const latex = latexOfEquation(node.attrs)
  return latex === '' ? t('document.math.placeholder') : latex
}

/** `false` when the MathML does not pass the filter. */
export function pastedMathAttrs(element: HTMLElement): Record<string, unknown> | false {
  return mathAttrsOfMarkup(new XMLSerializer().serializeToString(element))
}

/** Without OMML, with LaTeX derived from MathML. */
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

/** Also the equation editor preview. */
export function buildMath(element: MathElement, doc: Document): Element {
  const built = doc.createElementNS(MATHML_NAMESPACE, element.tag)
  for (const [name, value] of Object.entries(element.attrs)) built.setAttribute(name, value)
  for (const child of element.children) {
    built.append(typeof child === 'string' ? doc.createTextNode(child) : buildMath(child, doc))
  }
  return built
}

/**
 * Double click and Enter reach `EditEquation`, which decides whether to open for editing or
 * viewing. High priority, so Enter comes before the paragraph split, without touching the schema
 * order.
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
          // The signature is ProseMirror's, not ours.
          // eslint-disable-next-line max-params
          handleDoubleClickOn(view, _pos, node, nodePos, _event, direct) {
            if (!direct || node.type.name !== 'math') return false
            // The node position: on the right half, the click lands after the equation.
            view.dispatch(view.state.tr.setSelection(NodeSelection.create(view.state.doc, nodePos)))
            emitEditorCommand(EditorCommand.EditEquation)
            return true
          },
        },
      }),
    ]
  },
})
