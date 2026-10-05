import { useEffect, useMemo, useState } from 'react'
import { familiesInDocument, orderFontFamilies } from '@services/document/font-list.js'
import { useT } from '../../i18n.js'
import { useWorkspace } from '../../state/workspace.js'

/** The order comes from `font-list.ts`. Installed fonts are requested once per window. */
export function useFontFamilies(activeFamily: string): readonly { value: string; label: string }[] {
  const t = useT()
  const initialDoc = useWorkspace((state) => state.initialDoc)
  const [installed, setInstalled] = useState<readonly string[]>([])

  useEffect(() => {
    let alive = true

    void window.api.fonts.list({}).then((result) => {
      // No warning: the list is a convenience.
      if (alive && result.ok) setInstalled(result.data.families)
    })

    return () => {
      alive = false
    }
  }, [])

  // From the mount-time model, not the live content: walking the tree on every key press would be
  // costly.
  const inDocument = useMemo(() => familiesInDocument(initialDoc), [initialDoc])

  return useMemo(() => {
    const families = orderFontFamilies(installed, inDocument)

    // The cursor's font goes in even outside the three sources: otherwise the `<select>` would lie.
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
