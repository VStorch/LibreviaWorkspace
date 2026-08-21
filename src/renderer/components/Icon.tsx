/**
 * Ícones da barra de ferramentas.
 *
 * Desenhados aqui em vez de virem de uma biblioteca: são traços simples, e uma
 * dependência de ícones traria centenas de arquivos, um passo de build e mais
 * uma licença para auditar — por nada.
 *
 * Documento e planilha usam o mesmo conjunto. Antes a planilha desenhava seus
 * botões com caracteres soltos (`▧`, `❄`, `,00→`), o que fazia as duas metades
 * do programa parecerem dois programas.
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
  | 'text-color'
  | 'fill-color'
  | 'decimal-less'
  | 'decimal-more'
  | 'borders-all'
  | 'borders-none'
  | 'freeze'
  | 'unfreeze'
  | 'file-document'
  | 'file-spreadsheet'
  | 'folder-open'

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

  // A barrinha colorida abaixo destes dois é desenhada pelo CSS, com a cor
  // escolhida — o ícone mostra *o que* recebe a cor, a barra mostra *qual*.
  'text-color': <path d="M6 17 12 6l6 11M8.4 13.2h7.2" />,
  'fill-color': (
    <path d="M7.4 3.4 9.9 5.9M9.9 5.9 4.7 11.1a1.6 1.6 0 0 0 0 2.2l4.8 4.8a1.6 1.6 0 0 0 2.2 0l5.2-5.2zM20 8.6c0 1-.8 1.8-1.8 1.8s-1.8-.8-1.8-1.8 1.8-3.2 1.8-3.2 1.8 2.2 1.8 3.2z" />
  ),

  // Um dígito, a vírgula e a direção: para a direita ganha casa decimal, para a
  // esquerda perde. Desenhar ",00→" com traços viraria borrão a 18 pixels.
  'decimal-less': (
    <path d="M4.5 17.5h.01M9.5 11.5a2.2 3 0 1 0 0 6 2.2 3 0 0 0 0-6M19.5 14.5H15m2 2-2-2 2-2" />
  ),
  'decimal-more': (
    <path d="M4.5 17.5h.01M9.5 11.5a2.2 3 0 1 0 0 6 2.2 3 0 0 0 0-6M15 14.5h4.5m-2-2 2 2-2 2" />
  ),

  'borders-all': <path d="M4 4h16v16H4zM12 4v16M4 12h16" />,
  'borders-none': <path strokeDasharray="3 3" d="M4 4h16v16H4zM12 4v16M4 12h16" />,

  // Congelado é traço grosso, solto é tracejado: o mesmo par de grade, com a
  // linha de corte dizendo se ela prende ou não.
  freeze: (
    <>
      <path d="M4 4h16v16H4z" />
      <path d="M4 9.5h16M9.5 4v16" strokeWidth="2.6" />
    </>
  ),
  unfreeze: (
    <>
      <path d="M4 4h16v16H4z" />
      <path d="M4 9.5h16M9.5 4v16" strokeDasharray="2.5 2.5" />
    </>
  ),

  'file-document': <path d="M6 3h8l4 4v14H6zM14 3v4h4M9 12h6M9 16h4" />,
  'file-spreadsheet': <path d="M6 3h8l4 4v14H6zM14 3v4h4M8.5 11.5h7M8.5 15.5h7M12 11.5v8" />,
  'folder-open': <path d="M3.5 19V6H9l2 2.5h8.5V11M3.5 19l2.8-7.5h16.2L19.7 19z" />,
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
