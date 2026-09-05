import type { LossInventory } from '@shared/types.js'

/**
 * O que o inventário de um arquivo aberto manda a interface fazer.
 *
 * São duas decisões diferentes, e por isso duas funções: uma diz se há aviso a
 * mostrar, a outra se o documento abre travado. Misturá-las é como se acaba
 * avisando sempre e travando por qualquer coisa.
 */

/**
 * Há o que avisar?
 *
 * O aviso só aparece quando há: um alerta que abre em todo arquivo é um alerta
 * que o usuário fecha sem ler.
 */
export function hasReportableLoss(inventory: LossInventory | undefined): boolean {
  return inventory !== undefined && (inventory.invisible.length > 0 || inventory.lost.length > 0)
}

/**
 * Há recurso que some se o bloco que o ancora for editado?
 *
 * É o que decide o somente leitura. Perda de aparência — posicionamento de
 * imagem, decoração — não entra: quase todo documento do corpus tem alguma, e
 * travar por causa disso travaria o uso normal sem motivo.
 */
export function locksEditing(inventory: LossInventory | undefined): boolean {
  return inventory !== undefined && inventory.structural.length > 0
}
