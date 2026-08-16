#!/usr/bin/env node
/**
 * Build e portão de licenças do sidecar .NET.
 *
 * Existe como script Node, e não como linha de `npm scripts`, por dois motivos:
 * o Windows é plataforma-alvo (§8.1 do plano) e citação de shell quebra lá; e o
 * portão de licenças precisa de lógica de verdade, não de um `grep`.
 *
 *   node scripts/sidecar.mjs build [--all]   publica o binário
 *   node scripts/sidecar.mjs licenses        reprova licença fora da allowlist
 */

import { spawnSync } from 'node:child_process'
import { join, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'
import { ALLOWED_LICENSES, collectNuGetLicenses, nugetAssetsExist } from './nuget-licenses.mjs'

const root = join(dirname(fileURLToPath(import.meta.url)), '..')
const project = join(root, 'sidecar', 'src', 'Librevia.Format')

const RUNTIME_IDENTIFIERS = ['linux-x64', 'win-x64']

function currentRuntimeIdentifier() {
  if (process.arch !== 'x64') {
    fail(`sem alvo publicado para ${process.platform}-${process.arch}`)
  }
  if (process.platform === 'linux') return 'linux-x64'
  if (process.platform === 'win32') return 'win-x64'
  return fail(`sem alvo publicado para ${process.platform}`)
}

function fail(message) {
  console.error(`\n✖ ${message}\n`)
  process.exit(1)
}

function run(command, args) {
  const result = spawnSync(command, args, { stdio: 'inherit', shell: false })
  if (result.error) fail(`não foi possível executar ${command}: ${result.error.message}`)
  if (result.status !== 0) fail(`${command} terminou com código ${result.status}`)
}

function build(all) {
  const targets = all ? RUNTIME_IDENTIFIERS : [currentRuntimeIdentifier()]

  for (const rid of targets) {
    console.log(`\n▸ publicando sidecar para ${rid}`)
    run('dotnet', [
      'publish',
      project,
      '--configuration',
      'Release',
      '--runtime',
      rid,
      '--output',
      join(root, 'resources', 'sidecar', rid),
      // Sem restauração de rede depois da primeira: o projeto é offline.
      '--nologo',
    ])
  }
}

/**
 * Verifica o grafo NuGet resolvido.
 *
 * A regra dura: **ausência de expressão SPDX reprova.** Um pacote que só
 * informa `licenseUrl` está apontando para um texto que ninguém leu — foi assim
 * que a Six Labors trocou Apache-2.0 por licença própria a partir da versão 2
 * de `SixLabors.Fonts`, que é dependência transitiva do ClosedXML. Aceitar
 * "não declarou" como "deve estar tudo bem" seria abrir justamente o buraco que
 * este portão existe para fechar.
 */
function licenses() {
  if (!nugetAssetsExist()) {
    fail('project.assets.json ausente — rode `dotnet restore sidecar/Librevia.Format.slnx` antes')
  }

  const packages = collectNuGetLicenses()
  const problems = []
  const report = []

  for (const { name, version, license, found } of packages) {
    if (!found) {
      problems.push(`${name}/${version}: pacote não encontrado no cache local`)
      continue
    }

    report.push(`  ${`${name}/${version}`.padEnd(46)} ${license ?? '(não declarada)'}`)

    if (license === null) {
      problems.push(`${name}/${version}: não declara expressão de licença SPDX`)
    } else if (!ALLOWED_LICENSES.has(license)) {
      problems.push(`${name}/${version}: licença "${license}" fora da allowlist`)
    }
  }

  console.log(`\nLicenças NuGet (${packages.length} pacotes):`)
  console.log(report.join('\n'))

  if (problems.length > 0) {
    fail(`licenças reprovadas:\n  - ${problems.join('\n  - ')}`)
  }
  console.log('\n✓ todas dentro da allowlist\n')
}

const [command, ...flags] = process.argv.slice(2)

switch (command) {
  case 'build':
    build(flags.includes('--all'))
    break
  case 'licenses':
    licenses()
    break
  default:
    fail(`uso: node scripts/sidecar.mjs build [--all] | licenses`)
}
