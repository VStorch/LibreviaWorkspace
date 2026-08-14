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
import { readFileSync, existsSync, readdirSync } from 'node:fs'
import { join, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = join(dirname(fileURLToPath(import.meta.url)), '..')
const project = join(root, 'sidecar', 'src', 'Librevia.Format')
const assetsPath = join(project, 'obj', 'project.assets.json')

/** Mesma allowlist do lado npm — ver §4.4 do plano. */
const ALLOWED_LICENSES = new Set([
  'MIT',
  'MIT-0',
  'ISC',
  '0BSD',
  'BSD-2-Clause',
  'BSD-3-Clause',
  'Apache-2.0',
  'CC0-1.0',
  'Unlicense',
])

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
  if (!existsSync(assetsPath)) {
    fail('project.assets.json ausente — rode `dotnet restore sidecar/Librevia.Format.sln` antes')
  }

  const assets = JSON.parse(readFileSync(assetsPath, 'utf8'))
  const packageFolders = Object.keys(assets.packageFolders ?? {})
  if (packageFolders.length === 0) fail('nenhuma pasta de pacotes NuGet no project.assets.json')

  const resolved = new Set()
  for (const target of Object.values(assets.targets ?? {})) {
    for (const [key, entry] of Object.entries(target)) {
      if (entry.type === 'package') resolved.add(key)
    }
  }

  const problems = []
  const report = []

  for (const key of [...resolved].sort()) {
    const [name, version] = key.split('/')
    const nuspec = findNuspec(packageFolders, name, version)

    if (nuspec === null) {
      problems.push(`${key}: pacote não encontrado no cache local`)
      continue
    }

    const license = licenseOf(nuspec)
    report.push(`  ${key.padEnd(46)} ${license ?? '(não declarada)'}`)

    if (license === null) {
      problems.push(`${key}: não declara expressão de licença SPDX`)
    } else if (!ALLOWED_LICENSES.has(license)) {
      problems.push(`${key}: licença "${license}" fora da allowlist`)
    }
  }

  console.log(`\nLicenças NuGet (${resolved.size} pacotes):`)
  console.log(report.join('\n'))

  if (problems.length > 0) {
    fail(`licenças reprovadas:\n  - ${problems.join('\n  - ')}`)
  }
  console.log('\n✓ todas dentro da allowlist\n')
}

function findNuspec(packageFolders, name, version) {
  for (const folder of packageFolders) {
    // O cache do NuGet usa nomes em minúsculas.
    const directory = join(folder, name.toLowerCase(), version.toLowerCase())
    if (!existsSync(directory)) continue

    const file = readdirSync(directory).find((entry) => entry.toLowerCase().endsWith('.nuspec'))
    if (file !== undefined) return join(directory, file)
  }
  return null
}

function licenseOf(nuspecPath) {
  const xml = readFileSync(nuspecPath, 'utf8')
  const expression = /<license\s+type="expression"\s*>([^<]+)<\/license>/i.exec(xml)
  return expression?.[1]?.trim() ?? null
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
