import { useMemo, type ElementType, type ReactNode, type Ref } from 'react'
import { markdownToNodes } from '@services/document/markdown-import.js'
import type { DocumentNode } from '@services/document/model.js'
import { blockTagOf, linkOf, listStartOf, MARK_TAGS, type NodeMark } from './markdown-tags.js'

function withMark(content: ReactNode, mark: NodeMark): ReactNode {
  if (mark.type === 'link') {
    return (
      <span className="ai-preview__link" title={linkOf(mark)}>
        {content}
      </span>
    )
  }
  const Tag = MARK_TAGS[mark.type] as ElementType | undefined
  return Tag === undefined ? content : <Tag>{content}</Tag>
}

/** React elements with escaped text: the answer never reaches the page as HTML. */
function rendered(node: DocumentNode, key: number): ReactNode {
  if (node.type === 'text') {
    return <span key={key}>{(node.marks ?? []).reduce(withMark, node.text ?? '')}</span>
  }
  if (node.type === 'hardBreak') return <br key={key} />
  const Tag = blockTagOf(node) as ElementType
  const start = listStartOf(node)
  return (
    <Tag key={key} start={start}>
      {(node.content ?? []).map(rendered)}
    </Tag>
  )
}

export function MarkdownBlocks({ markdown }: { readonly markdown: string }): React.JSX.Element {
  const nodes = useMemo(() => markdownToNodes(markdown), [markdown])
  return <>{nodes.map(rendered)}</>
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
  return (
    <div ref={ref} className="ai-pane__output ai-preview" role="region" aria-label={label} tabIndex={0}>
      <MarkdownBlocks markdown={markdown} />
    </div>
  )
}
