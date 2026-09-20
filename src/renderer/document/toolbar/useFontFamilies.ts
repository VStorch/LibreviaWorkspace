import { useEffect, useMemo, useState } from 'react'
import { familiesInDocument, orderFontFamilies } from '@services/document/font-list.js'
import { useWorkspace } from '../../state/workspace.js'

/**
 * As fontes que o seletor da barra oferece.
 *
 * Três origens, nesta ordem de importância: as do documento aberto, as que o
 * instalador garante, e as instaladas na máquina. A ordem e a fusão são de
 * `@services/document/font-list.ts`; aqui fica só o que é de React — quando
 * perguntar ao main e o que fazer enquanto a resposta não chega.
 *
 * A lista das instaladas é pedida **uma vez por janela**. É informação do
 * ambiente, não do documento: repetir a pergunta a cada arquivo aberto gastaria
 * um processo do sistema por nada.
 */
export function useFontFamilies(activeFamily: string): readonly { value: string; label: string }[] {
  const initialDoc = useWorkspace((state) => state.initialDoc)
  const [installed, setInstalled] = useState<readonly string[]>([])

  useEffect(() => {
    let alive = true

    void window.api.fonts.list({}).then((result) => {
      // Falha aqui não vira aviso na tela: a lista é um conforto, e sem ela a
      // barra segue oferecendo as fontes que viajam no instalador. Interromper o
      // trabalho de quem escreve por causa de um `fc-list` ausente seria pior do
      // que a lista curta.
      if (alive && result.ok) setInstalled(result.data.families)
    })

    return () => {
      alive = false
    }
  }, [])

  // As do documento vêm do modelo com que o editor foi montado, e não do
  // conteúdo ao vivo: varrer a árvore inteira a cada tecla digitada custaria
  // mais do que tudo o que esta barra faz junto.
  const inDocument = useMemo(() => familiesInDocument(initialDoc), [initialDoc])

  return useMemo(() => {
    const families = orderFontFamilies(installed, inDocument)

    // A fonte do cursor entra na lista mesmo quando não está em nenhuma das três
    // origens — texto colado de outro documento traz nome que ninguém listou, e
    // um `<select>` sem a opção do seu próprio valor mostra a primeira da lista:
    // a barra passaria a mentir sobre o que está debaixo do cursor.
    if (
      activeFamily !== '' &&
      !families.some((family) => family.toLowerCase() === activeFamily.toLowerCase())
    ) {
      families.unshift(activeFamily)
    }

    return [
      { value: '', label: 'Fonte padrão' },
      ...families.map((family) => ({ value: family, label: family })),
    ]
  }, [installed, inDocument, activeFamily])
}
