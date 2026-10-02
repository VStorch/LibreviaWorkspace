import { Node } from '@tiptap/core'
import type { Node as ProseMirrorNode } from '@tiptap/pm/model'
import { MATHML_NAMESPACE, sanitizeMathMl, type MathElement } from '@services/document/mathml.js'
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
 * edição, que esta fase ainda não tem.
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
      omml: { default: null, parseHTML: (element) => element.getAttribute('data-omml') },
      mathml: { default: '', parseHTML: (element) => element.getAttribute('data-mathml') ?? '' },
      latex: { default: '', parseHTML: (element) => element.getAttribute('data-latex') ?? '' },
      display: { default: false, parseHTML: (element) => element.getAttribute('data-display') === 'true' },
      jc: { default: null, parseHTML: (element) => element.getAttribute('data-jc') },
      lossy: { default: [], parseHTML: (element) => lossyOf(element.getAttribute('data-lossy')) },
      editable: { default: true, parseHTML: (element) => element.getAttribute('data-editable') !== 'false' },
    }
  },

  parseHTML() {
    return [{ tag: 'span[data-math]' }]
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

/** O texto que a equação deixa ao ser copiada: o LaTeX, quando há, ou o marcador. */
export function plainTextOf(node: ProseMirrorNode): string {
  const latex = String(node.attrs['latex'] ?? '')
  return latex === '' ? t('document.math.placeholder') : latex
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

function buildMath(element: MathElement, doc: Document): Element {
  const built = doc.createElementNS(MATHML_NAMESPACE, element.tag)
  for (const [name, value] of Object.entries(element.attrs)) built.setAttribute(name, value)
  for (const child of element.children) {
    built.append(typeof child === 'string' ? doc.createTextNode(child) : buildMath(child, doc))
  }
  return built
}
