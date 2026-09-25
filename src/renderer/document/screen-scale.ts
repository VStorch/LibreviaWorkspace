/**
 * Quantos pixels de tela vale um pixel de CSS dentro da folha.
 *
 * O zoom da folha é um `transform` sobre a pilha (`DocumentEditor`): o layout
 * continua em 100 %, e só o que vem da tela — `clientX`, `getBoundingClientRect`
 * — chega multiplicado. Quem converte gesto em medida do documento divide por
 * isto; sem a divisão, arrastar a alça da imagem a 150 % gravava a imagem com
 * uma vez e meia o tamanho que a pessoa viu.
 *
 * Lido do próprio elemento, e não da preferência: é a razão entre o que a tela
 * desenha e o que o layout mede, e vale para qualquer transformação no caminho.
 */
export function screenScaleOf(element: HTMLElement | null): number {
  if (element === null || element.offsetWidth <= 0) return 1
  const width = element.getBoundingClientRect().width
  return width > 0 ? width / element.offsetWidth : 1
}
