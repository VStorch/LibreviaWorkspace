/**
 * Ícones da barra de ferramentas.
 *
 * Desenhados aqui em vez de virem de uma biblioteca: são dezoito traços
 * simples, e uma dependência de ícones traria centenas de arquivos, um passo
 * de build e mais uma licença para auditar — por nada.
 */

export type IconName =
  | 'bold'
  | 'italic'
  | 'underline'
  | 'strike'
  | 'align-left'
  | 'align-center'
  | 'align-right'
  | 'align-justify'
  | 'bullet-list'
  | 'ordered-list'
  | 'indent'
  | 'outdent'
  | 'table'
  | 'image'
  | 'link'
  | 'page-break'
  | 'search'
  | 'page-setup'
  | 'print-preview'

const PATHS: Record<IconName, React.ReactNode> = {
  bold: <path d="M7 5h5.5a3.5 3.5 0 0 1 0 7H7zm0 7h6.5a3.5 3.5 0 0 1 0 7H7z" />,
  italic: <path d="M15 5h-4M13 19H9M14 5l-3 14" />,
  underline: <path d="M7 4v6a5 5 0 0 0 10 0V4M5 20h14" />,
  strike: (
    <path d="M5 12h14M8 8a3.5 3.5 0 0 1 3.5-3h1A3.5 3.5 0 0 1 16 8M8 16a3.5 3.5 0 0 0 3.5 3h1a3.5 3.5 0 0 0 3.5-3" />
  ),
  'align-left': <path d="M4 6h16M4 10h10M4 14h16M4 18h10" />,
  'align-center': <path d="M4 6h16M7 10h10M4 14h16M7 18h10" />,
  'align-right': <path d="M4 6h16M10 10h10M4 14h16M10 18h10" />,
  'align-justify': <path d="M4 6h16M4 10h16M4 14h16M4 18h16" />,
  'bullet-list': <path d="M9 6h11M9 12h11M9 18h11M4.5 6h.01M4.5 12h.01M4.5 18h.01" />,
  'ordered-list': <path d="M10 6h10M10 12h10M10 18h10M4 5h1v4M4 15h2v1H4v2h2" />,
  indent: <path d="M4 6h16M10 10h10M10 14h10M4 18h16M4 10l3 2-3 2z" />,
  outdent: <path d="M4 6h16M10 10h10M10 14h10M4 18h16M7 10l-3 2 3 2z" />,
  table: <path d="M4 5h16v14H4zM4 10h16M4 15h16M10 5v14" />,
  image: <path d="M4 5h16v14H4zM4 15l4.5-4.5 4 4L16 11l4 4" />,
  link: (
    <path d="M10 13a4 4 0 0 0 5.7.3l2.6-2.6a4 4 0 0 0-5.7-5.7l-1.5 1.5M14 11a4 4 0 0 0-5.7-.3l-2.6 2.6a4 4 0 0 0 5.7 5.7l1.5-1.5" />
  ),
  'page-break': <path d="M6 4h12M6 20h12M3 12h4M10 12h4M17 12h4M8 7l4-3 4 3M8 17l4 3 4-3" />,
  search: <path d="M11 4a7 7 0 1 0 0 14 7 7 0 0 0 0-14zM16 16l4 4" />,
  'page-setup': <path d="M5 3h9l5 5v13H5zM14 3v5h5M8 13h8M8 17h5" />,
  'print-preview': <path d="M7 9V4h10v5M7 18H5v-6h14v6h-2M8 15h8v6H8z" />,
}

export function Icon({ name }: { readonly name: IconName }): React.JSX.Element {
  return (
    <svg
      className="icon"
      viewBox="0 0 24 24"
      width="18"
      height="18"
      fill="none"
      stroke="currentColor"
      strokeWidth="1.7"
      strokeLinecap="round"
      strokeLinejoin="round"
      aria-hidden="true"
    >
      {PATHS[name]}
    </svg>
  )
}
