import { useEffect, useRef, useState } from 'react'
import { pxToMm, type PageSetup } from '@services/document/model.js'
import { type BandHeights } from '@services/document/band.js'

/**
 * A altura desenhada do cabeçalho e do rodapé, seção por seção.
 *
 * É a única parte da conta de margem que nenhum arquivo diz: um cabeçalho em
 * grade ocupa o que a fonte e a quebra derem, e isso só existe depois de
 * desenhar. Medido na primeira folha de cada seção — as outras da mesma seção
 * repetem a mesma faixa — e arredondado a um décimo de milímetro, porque a
 * medida do navegador oscila sozinha e cada oscilação repaginaria o documento
 * inteiro.
 *
 * Não há laço: a altura da faixa não depende de onde o texto caiu. A seção que
 * ainda não tem folha desenhada fica com a altura da anterior até ter uma.
 *
 * @param sheets Que seção abre cada folha — muda quando uma seção ganha ou perde
 * folhas, e aí há faixa nova para medir.
 */
export function useBandHeights(sections: readonly PageSetup[], revision: number, sheets = ''): BandHeights[] {
  const [bands, setBands] = useState<BandHeights[]>([])
  const last = useRef<BandHeights[]>([])

  useEffect(() => {
    // Sem chamar o `setState` quando nada mudou: o efeito roda a cada tecla, e
    // um estado pedido a cada desenho — mesmo igual — encadeava desenhos na
    // digitação rápida até o React desistir (erro 185).
    const measure = (): void => {
      const next = measureBands(sections.length)
      if (sameBands(last.current, next)) return
      last.current = next
      setBands(next)
    }

    measure()

    const layers = document.querySelectorAll('.paper-bands')
    if (layers.length === 0) return undefined

    const observer = new ResizeObserver(measure)
    for (const layer of firstOfEachSection(layers)) {
      observer.observe(layer)
      for (const band of layer.querySelectorAll('.band')) observer.observe(band)
    }
    return () => observer.disconnect()
  }, [sections, revision, sheets])

  return bands
}

/** A primeira camada de faixas de cada seção. */
function firstOfEachSection(layers: NodeListOf<Element>): Element[] {
  const seen = new Set<string>()
  const first: Element[] = []
  for (const layer of layers) {
    const section = layer.getAttribute('data-section') ?? '0'
    if (seen.has(section)) continue
    seen.add(section)
    first.push(layer)
  }
  return first
}

function measureBands(count: number): BandHeights[] {
  const layers = firstOfEachSection(document.querySelectorAll('.paper-bands'))
  const bySection = new Map(layers.map((layer) => [Number(layer.getAttribute('data-section') ?? 0), layer]))

  const heightOf = (layer: Element | undefined, kind: string): number => {
    const band = layer?.querySelector(`.band--${kind}`)
    if (band === null || band === undefined) return 0
    return Math.round(pxToMm((band as HTMLElement).offsetHeight) * 10) / 10
  }

  const result: BandHeights[] = []
  for (let section = 0; section < Math.max(count, 1); section++) {
    const layer = bySection.get(section)
    result.push(
      layer === undefined
        ? (result.at(-1) ?? { headerMm: 0, footerMm: 0 })
        : { headerMm: heightOf(layer, 'header'), footerMm: heightOf(layer, 'footer') },
    )
  }
  return result
}

function sameBands(left: readonly BandHeights[], right: readonly BandHeights[]): boolean {
  return (
    left.length === right.length &&
    left.every(
      (band, index) => band.headerMm === right[index]!.headerMm && band.footerMm === right[index]!.footerMm,
    )
  )
}
