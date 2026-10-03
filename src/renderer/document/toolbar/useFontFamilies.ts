import { useEffect, useMemo, useState } from 'react'
import { familiesInDocument, orderFontFamilies } from '@services/document/font-list.js'
import { useT } from '../../i18n.js'
import { useWorkspace } from '../../state/workspace.js'

/** A ordem é de `font-list.ts`. As instaladas são pedidas uma vez por janela. */
export function useFontFamilies(activeFamily: string): readonly { value: string; label: string }[] {
  const t = useT()
  const initialDoc = useWorkspace((state) => state.initialDoc)
  const [installed, setInstalled] = useState<readonly string[]>([])

  useEffect(() => {
    let alive = true

    void window.api.fonts.list({}).then((result) => {
      // Sem aviso: a lista é um conforto.
      if (alive && result.ok) setInstalled(result.data.families)
    })

    return () => {
      alive = false
    }
  }, [])

  // Do modelo de montagem, e não do conteúdo ao vivo: varrer a árvore a cada tecla custaria caro.
  const inDocument = useMemo(() => familiesInDocument(initialDoc), [initialDoc])

  return useMemo(() => {
    const families = orderFontFamilies(installed, inDocument)

    // A fonte do cursor entra mesmo fora das três origens: o `<select>` mentiria.
    if (
      activeFamily !== '' &&
      !families.some((family) => family.toLowerCase() === activeFamily.toLowerCase())
    ) {
      families.unshift(activeFamily)
    }

    return [
      { value: '', label: t('document.styleAndFont.defaultFont') },
      ...families.map((family) => ({ value: family, label: family })),
    ]
  }, [installed, inDocument, activeFamily, t])
}
