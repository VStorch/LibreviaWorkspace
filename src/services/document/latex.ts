import temml from 'temml'
import {
  mathMlToString,
  parseMathMl,
  sanitizeMathMl,
  MATH_BOX_CLASS,
  type MathChild,
  type MathElement,
} from './mathml.js'

/**
 * Temml writes MathML Core and runs offline. Three of its outputs are translated before the
 * `mathml.ts` filter: the `menclose` of `\overline` and `\underline` becomes `mover`/`munder`
 * (`m:bar`); `\boxed` becomes the `mrow` with the box class; and the prescript `{}_a^b X` becomes
 * `mmultiscripts` (`m:sPre`). The source lives in the node's `latex` attribute, so `semantics` is
 * unwrapped.
 */

export type LatexResult =
  | { readonly ok: true; readonly mathml: string; readonly tree: MathElement }
  | { readonly ok: false; readonly error: string }

export function latexToMathMl(latex: string, display: boolean): LatexResult {
  let raw: string
  try {
    raw = temml.renderToString(latex, {
      displayMode: display,
      throwOnError: true,
      annotate: false,
      trust: false,
    })
  } catch (error) {
    // Besides Temml's `ParseError`, a `TypeError` on `a^` is also LaTeX that does not close.
    const message = error instanceof Error ? error.message : String(error)
    return { ok: false, error: message.replace(/^ParseError:\s*/, '').trim() }
  }

  const parsed = parseMathMl(raw)
  if (parsed === null) return { ok: false, error: 'MathML' }
  const tree = sanitizeMathMl(mathMlToString(normalize(parsed)))
  if (tree === null) return { ok: false, error: 'MathML' }
  return { ok: true, mathml: mathMlToString(tree), tree }
}

function element(tag: string, attrs: Record<string, string>, children: readonly MathChild[]): MathElement {
  return { tag, attrs, children }
}

/** Temml's MathML in the vocabulary the filter and `OmmlMath.cs` know. */
export function normalize(node: MathElement): MathElement {
  const children = scripts(
    node.children.map((child) => (typeof child === 'string' ? child : normalize(child))).flatMap(unwrap),
  )

  if (node.tag === 'menclose') {
    const notation = node.attrs['notation'] ?? ''
    const body = element('mrow', {}, children)
    if (notation === 'top') {
      return element('mover', { accent: 'true' }, [body, element('mo', { stretchy: 'true' }, ['‾'])])
    }
    if (notation === 'bottom') {
      return element('munder', { accentunder: 'true' }, [body, element('mo', { stretchy: 'true' }, ['‾'])])
    }
    if (notation === 'box') return element('mrow', { class: MATH_BOX_CLASS }, children)
    return body
  }

  if (node.tag === 'mrow' && /border\s*:/.test(node.attrs['style'] ?? '')) {
    return element('mrow', { class: MATH_BOX_CLASS }, children)
  }

  return { ...node, children }
}

function unwrap(child: MathChild): MathChild[] {
  if (typeof child === 'string') return [child]
  if (child.tag === 'annotation' || child.tag === 'annotation-xml') return []
  if (child.tag === 'semantics') {
    const first = child.children.find((item): item is MathElement => typeof item !== 'string')
    return first === undefined ? [] : [first]
  }
  return [child]
}

/** `{}_a^b X`, the empty-base script before the base, into `mmultiscripts`. */
function scripts(children: MathChild[]): MathChild[] {
  const result: MathChild[] = []
  for (let i = 0; i < children.length; i++) {
    const child = children[i]!
    const next = children[i + 1]
    if (
      typeof child !== 'string' &&
      (child.tag === 'msub' || child.tag === 'msup' || child.tag === 'msubsup') &&
      emptyBase(child) &&
      next !== undefined &&
      typeof next !== 'string'
    ) {
      const [, first, second] = child.children
      const none = element('none', {}, [])
      const sub = child.tag === 'msup' ? none : (first ?? none)
      const sup = child.tag === 'msub' ? none : child.tag === 'msup' ? (first ?? none) : (second ?? none)
      result.push(element('mmultiscripts', {}, [next, element('mprescripts', {}, []), sub, sup]))
      i++
      continue
    }
    result.push(child)
  }
  return result
}

function emptyBase(script: MathElement): boolean {
  const base = script.children[0]
  return base !== undefined && typeof base !== 'string' && base.tag === 'mrow' && base.children.length === 0
}
