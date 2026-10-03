/**
 * O zoom é um `transform`: só o que vem da tela chega multiplicado. Quem converte
 * gesto em medida divide por isto. Lido do elemento, para valer com qualquer
 * transformação no caminho.
 */
export function screenScaleOf(element: HTMLElement | null): number {
  if (element === null || element.offsetWidth <= 0) return 1
  const width = element.getBoundingClientRect().width
  return width > 0 ? width / element.offsetWidth : 1
}
