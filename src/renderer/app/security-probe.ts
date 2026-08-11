/**
 * Sonda de isolamento.
 *
 * O critério de aceite da Fase 0 é que o renderer *comprovadamente* não alcance
 * o Node.js. O teste unitário verifica a configuração; esta sonda verifica o
 * comportamento real, dentro do Chromium, e mostra o resultado na tela — porque
 * configuração correta e efeito correto não são a mesma coisa.
 */

export interface ProbeResult {
  readonly label: string
  readonly expectation: string
  readonly passed: boolean
}

/** `typeof` em identificador não declarado não lança — por isso a checagem é assim. */
function isUndefinedGlobal(name: string): boolean {
  return (globalThis as Record<string, unknown>)[name] === undefined
}

export function runSecurityProbes(): readonly ProbeResult[] {
  return [
    {
      label: 'require',
      expectation: 'indisponível',
      passed: isUndefinedGlobal('require'),
    },
    {
      label: 'process',
      expectation: 'indisponível',
      passed: isUndefinedGlobal('process'),
    },
    {
      label: 'module',
      expectation: 'indisponível',
      passed: isUndefinedGlobal('module'),
    },
    {
      label: 'Buffer',
      expectation: 'indisponível',
      passed: isUndefinedGlobal('Buffer'),
    },
    {
      label: 'window.api',
      expectation: 'exposto pelo contextBridge',
      passed: typeof window.api?.app?.ping === 'function',
    },
  ]
}
