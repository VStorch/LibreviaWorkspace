import { useMemo, type ElementType, type ReactNode, type Ref } from 'react'
import { markdownToNodes } from '@services/document/markdown-import.js'
import type { DocumentNode } from '@services/document/model.js'

type NodeMark = NonNullable<DocumentNode['marks']>[number]

const BLOCK_TAGS: Readonly<Record<string, ElementType>> = {
  paragraph: 'p',
  blockquote: 'blockquote',
  bulletList: 'ul',
  orderedList: 'ol',
  listItem: 'li',
  codeBlock: 'pre',
}

const MARK_TAGS: Readonly<Record<string, ElementType>> = {
  bold: 'strong',
  italic: 'em',
  strike: 's',
  code: 'code',
}

const HEADING_TAGS: readonly ElementType[] = ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']

function withMark(content: ReactNode, mark: NodeMark): ReactNode {
  if (mark.type === 'link') {
    return (
      <span className="ai-preview__link" title={String(mark.attrs?.['href'] ?? '')}>
        {content}
      </span>
    )
  }
  const Tag = MARK_TAGS[mark.type]
  return Tag === undefined ? content : <Tag>{content}</Tag>
}

function tagOf(node: DocumentNode): ElementType {
  if (node.type !== 'heading') return BLOCK_TAGS[node.type] ?? 'p'
  const level = Number(node.attrs?.['level']) || 1
  return HEADING_TAGS[Math.min(HEADING_TAGS.length, Math.max(1, level)) - 1]!
}

/** React elements with escaped text: the answer never reaches the page as HTML. */
function rendered(node: DocumentNode, key: number): ReactNode {
  if (node.type === 'text') {
    return <span key={key}>{(node.marks ?? []).reduce(withMark, node.text ?? '')}</span>
  }
  if (node.type === 'hardBreak') return <br key={key} />
  const Tag = tagOf(node)
  const start = node.type === 'orderedList' ? Number(node.attrs?.['start'] ?? 1) : undefined
  return (
    <Tag key={key} start={start}>
      {(node.content ?? []).map(rendered)}
    </Tag>
  )
}

export function AiMarkdownPreview({
  markdown,
  label,
  ref,
}: {
  readonly markdown: string
  readonly label: string
  readonly ref?: Ref<HTMLDivElement>
}): React.JSX.Element {
  const nodes = useMemo(() => markdownToNodes(markdown), [markdown])
  return (
    <div ref={ref} className="ai-pane__output ai-preview" role="region" aria-label={label} tabIndex={0}>
      {nodes.map(rendered)}
    </div>
  )
}
