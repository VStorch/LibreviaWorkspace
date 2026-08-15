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

      const found: number[] = []
      let pageEnd = pageHeight

      for (const block of Array.from(element.children)) {
        const node = block as HTMLElement
        const bottom = node.offsetTop + node.offsetHeight

        // Quebra antes do bloco que estouraria a página — a mesma decisão que
        // o navegador toma ao imprimir, e o motivo de a linha não cortar texto.
        while (bottom > pageEnd) {
          const at = node.offsetTop > pageEnd - pageHeight ? node.offsetTop : pageEnd
          found.push(at)
          pageEnd = at + pageHeight

          // Bloco mais alto que uma página inteira (uma captura de tela grande):
          // não há onde quebrar, então segue.
          if (found.length > 500) break
        }
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
