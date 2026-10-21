#!/usr/bin/env node
/**
 * Build and license gate for the .NET sidecar, in Node because shell quoting breaks on Windows and
 * the gate needs logic.
 *
 *   node scripts/sidecar.mjs build [--all]   publishes the binary
 *   node scripts/sidecar.mjs licenses        fails on a license outside the allowlist
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
    fail(`no publish target for ${process.platform}-${process.arch}`)
  }
  if (process.platform === 'linux') return 'linux-x64'
  if (process.platform === 'win32') return 'win-x64'
  return fail(`no publish target for ${process.platform}`)
}

function fail(message) {
  console.error(`\n✖ ${message}\n`)
  process.exit(1)
}

function run(command, args) {
  const result = spawnSync(command, args, { stdio: 'inherit', shell: false })
  if (result.error) fail(`could not run ${command}: ${result.error.message}`)
  if (result.status !== 0) fail(`${command} exited with code ${result.status}`)
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
      // No network restore after the first: the project is offline.
      '--nologo',
    ])
  }
}

/**
 * Checks the resolved NuGet graph: without an SPDX expression, it fails. `licenseUrl` points to
 * text that changes, like the license of `SixLabors.Fonts`, a ClosedXML dependency.
 */
function licenses() {
  if (!nugetAssetsExist()) {
    fail('project.assets.json missing: run `dotnet restore sidecar/Librevia.Format.slnx` first')
  }

  const packages = collectNuGetLicenses()
  const problems = []
  const report = []

  for (const { name, version, license, found } of packages) {
    if (!found) {
      problems.push(`${name}/${version}: package not found in the local cache`)
      continue
    }

    report.push(`  ${`${name}/${version}`.padEnd(46)} ${license ?? '(not declared)'}`)

    if (license === null) {
      problems.push(`${name}/${version}: declares no SPDX license expression`)
    } else if (!ALLOWED_LICENSES.has(license)) {
      problems.push(`${name}/${version}: license "${license}" is outside the allowlist`)
    }
  }

  console.log(`\nNuGet licenses (${packages.length} packages):`)
  console.log(report.join('\n'))

  if (problems.length > 0) {
    fail(`rejected licenses:\n  - ${problems.join('\n  - ')}`)
  }
  console.log('\n✓ all within the allowlist\n')
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
