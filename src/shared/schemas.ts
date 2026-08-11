import { z } from 'zod'

/**
 * Schemas usados em mais de um lugar.
 *
 * A configuração de página é validada tanto ao ler um `.sdoc` do disco quanto
 * ao receber um pedido de impressão do renderer. Duas definições divergiriam,
 * e a divergência apareceria como margem errada no papel.
 */
export const pageSetupSchema = z.object({
  size: z.enum(['A4', 'Letter']),
  orientation: z.enum(['portrait', 'landscape']),
  margins: z.object({
    top: z.number(),
    right: z.number(),
    bottom: z.number(),
    left: z.number(),
  }),
  // Acrescentados na Fase 3. Opcionais para que documentos gravados antes
  // continuem abrindo — acréscimo compatível não exige nova versão de formato.
  header: z.string().max(500).default(''),
  footer: z.string().max(500).default(''),
})
