import { useEffect, useState } from 'react'
import { pxToMm, type PageSetup } from '@services/document/model.js'
import { type BandHeights } from '@services/document/band.js'

/**
 * A altura desenhada do cabeçalho e do rodapé.
 *
 * É a única parte da conta de margem que nenhum arquivo diz: um cabeçalho em
 * grade ocupa o que a fonte e a quebra derem, e isso só existe depois de
 * desenhar. Medido na primeira folha — as outras repetem a mesma faixa — e
 * arredondado a um décimo de milímetro, porque a medida do navegador oscila
 * sozinha e cada oscilação repaginaria o documento inteiro.
 *
 * Não há laço: a altura da faixa não depende de onde o texto caiu.
 */
export function useBandHeights(page: PageSetup, revision: number): BandHeights {
  const [bands, setBands] = useState<BandHeights>({ headerMm: 0, footerMm: 0 })

  useEffect(() => {
    const measure = (): void => {
      const next = measureBands()
      setBands((current) =>
        current.headerMm === next.headerMm && current.footerMm === next.footerMm ? current : next,
      )
    }

    measure()

    const sheet = document.querySelector('.paper-bands')
    if (sheet === null) return undefined

    const observer = new ResizeObserver(measure)
    observer.observe(sheet)
    for (const band of sheet.querySelectorAll('.band')) observer.observe(band)
    return () => observer.disconnect()
  }, [page, revision])

  return bands
}

function measureBands(): BandHeights {
  const sheet = document.querySelector('.paper-bands')

  const heightOf = (kind: string): number => {
    const band = sheet?.querySelector(`.band--${kind}`)
    if (band === null || band === undefined) return 0
    return Math.round(pxToMm((band as HTMLElement).offsetHeight) * 10) / 10
  }

  return { headerMm: heightOf('header'), footerMm: heightOf('footer') }
}
