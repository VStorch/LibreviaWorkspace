import { defineConfig } from '@playwright/test'

/**
 * Testes de ponta a ponta contra o aplicativo Electron de verdade.
 *
 * Eles rodam sobre `out/`, e não sobre o código-fonte: o que precisa ser
 * provado é o aplicativo empacotado — com preload sandboxed, contextIsolation e
 * o sidecar .NET publicado. Um teste que rodasse sobre o fonte não passaria por
 * nenhuma dessas fronteiras, que são justamente onde os erros desta arquitetura
 * aparecem.
 *
 * Um trabalhador só, e sem paralelismo: cada teste sobe um Electron inteiro, e
 * dois ao mesmo tempo disputariam o mesmo `userData` e a mesma trava de
 * instância única.
 */
export default defineConfig({
  testDir: './e2e',
  fullyParallel: false,
  workers: 1,
  // Subir o Electron e o sidecar leva alguns segundos; o padrão de 30 s deixa
  // pouca margem numa máquina carregada.
  timeout: 90_000,
  expect: { timeout: 15_000 },
  reporter: [['list']],
  // Falha em CI se alguém esquecer um `.only` — o resto da suíte não rodaria.
  forbidOnly: Boolean(process.env.CI),
})
