import { useEffect, useState, type RefObject } from 'react'
import { contentHeightMm, mmToPx, type PageSetup } from '@services/document/model.js'

/**
 * Marcas de onde a página termina.
 *
 * O editor **não pagina ao vivo** (docs/00-plano-tecnico.md §6.3): o texto rola
 * de forma contínua e a paginação real acontece na exportação para PDF. Sem
 * nenhuma marca, porém, o documento parece uma folha só, sem fim — e a pessoa
 * perde a noção de tamanho que o papel dá.
 *
 * Estas marcas são **medidas, não calculadas**: percorremos os blocos já
 * renderizados e quebramos antes do primeiro que não caberia. Por isso a linha
 * nunca corta um parágrafo ao meio. Ainda é uma estimativa — o Chromium aplica
 * regras de viúvas, órfãs e "manter junto" que só existem na hora de imprimir —
 * e por isso a marca diz "aproximada".
 */
export function usePageBreaks(
  pageRef: RefObject<HTMLElement | null>,
  page: PageSetup,
  revision: number,
): number[] {
  const [offsets, setOffsets] = useState<number[]>([])

  useEffect(() => {
    const element = pageRef.current?.querySelector<HTMLElement>('.page__content') ?? null
    if (element === null) return undefined

    const measure = (): void => {
      const pageHeight = mmToPx(contentHeightMm(page))
      if (pageHeight <= 0) {
        setOffsets([])
        return
      }

      const blocks = Array.from(element.children) as HTMLElement[]
      const found: number[] = []
      let pageStart = 0
      let index = 0

      while (index < blocks.length && found.length < 500) {
        const node = blocks[index]!
        const bottom = node.offsetTop + node.offsetHeight

        if (bottom - pageStart <= pageHeight) {
          index += 1
          continue
        }

        // Quebra antes do bloco que estouraria a página — a mesma decisão que
        // o navegador toma ao imprimir, e o motivo de a linha não cortar texto.
        let breakAt = node.offsetTop

        // Um título sozinho no pé da página desce junto com o que ele
        // apresenta. É a regra `break-after: avoid` que a exportação aplica; sem
        // ela, a marca cai um bloco depois de onde o PDF vai quebrar.
        // Sobe enquanto o bloco anterior pedir para ficar com o seguinte.
        let candidate = index
        while (candidate > 0) {
          const previous = blocks[candidate - 1]
          if (previous === undefined || !keepsWithNext(previous)) break
          if (previous.offsetTop <= pageStart) break
          candidate -= 1
          breakAt = previous.offsetTop
        }

        if (breakAt <= pageStart) {
          // Bloco mais alto que uma página inteira — uma captura de tela grande.
          // Não há onde quebrar sem cortá-lo, então ele transborda.
          pageStart = bottom
          index += 1
          continue
        }

        found.push(breakAt)
        pageStart = breakAt
        // `index` não avança: o mesmo bloco é reavaliado na página nova.
      }

      setOffsets(found)
    }

    measure()

    // O conteúdo muda de altura ao digitar, ao carregar imagem e ao redimensionar.
    const observer = new ResizeObserver(measure)
    observer.observe(element)
    return () => observer.disconnect()
  }, [pageRef, page, revision])

  return offsets
}

/**
 * O bloco pede para não ficar sozinho no pé da página.
 *
 * Vem de `w:keepNext` do documento; títulos entram por padrão, porque a folha
 * de impressão já lhes aplica `break-after: avoid`.
 */
const keepsWithNext = (element: HTMLElement): boolean =>
  element.hasAttribute('data-keep-next') || /^H[1-6]$/.test(element.tagName)

export function PageGuides({ offsets, topOffsetPx }: { offsets: number[]; topOffsetPx: number }) {
  if (offsets.length === 0) return null

  return (
    <>
      {offsets.map((offset, index) => (
        <div
          key={offset}
          className="page-guide"
          style={{ top: `${topOffsetPx + offset}px` }}
          aria-hidden="true"
        >
          <span className="page-guide__label">Página {index + 2}</span>
        </div>
      ))}
    </>
  )
}
