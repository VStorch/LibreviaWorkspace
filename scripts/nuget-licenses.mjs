/**
 * Licenses of the resolved NuGet packages, read from the local cache, so the gate (`sidecar.mjs
 * licenses`) and the notices (`notices.mjs`) do not diverge.
 */

import { existsSync, readFileSync, readdirSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = join(dirname(fileURLToPath(import.meta.url)), '..')
const assetsPath = join(root, 'sidecar', 'src', 'Librevia.Format', 'obj', 'project.assets.json')

/** A mesma allowlist do lado npm. */
export const ALLOWED_LICENSES = new Set([
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

export function nugetAssetsExist() {
  return existsSync(assetsPath)
}

/** `[{ name, version, license }]`, sorted. `license` is `null` when undeclared. */
export function collectNuGetLicenses() {
  const assets = JSON.parse(readFileSync(assetsPath, 'utf8'))
  const packageFolders = Object.keys(assets.packageFolders ?? {})
  if (packageFolders.length === 0) throw new Error('nenhuma pasta de pacotes NuGet no project.assets.json')

  const resolved = new Set()
  for (const target of Object.values(assets.targets ?? {})) {
    for (const [key, entry] of Object.entries(target)) {
      if (entry.type === 'package') resolved.add(key)
    }
  }

  return [...resolved].sort().map((key) => {
    const [name, version] = key.split('/')
    const nuspec = findNuspec(packageFolders, name, version)
    return {
      name,
      version,
      license: nuspec === null ? null : licenseOf(nuspec),
      found: nuspec !== null,
    }
  })
}

function findNuspec(packageFolders, name, version) {
  for (const folder of packageFolders) {
    // The NuGet cache uses lowercase names.
    const directory = join(folder, name.toLowerCase(), version.toLowerCase())
    if (!existsSync(directory)) continue

    const file = readdirSync(directory).find((entry) => entry.toLowerCase().endsWith('.nuspec'))
    if (file !== undefined) return join(directory, file)
  }
  return null
}

/**
 * Only an SPDX expression counts: `SixLabors.Fonts` is Apache-2.0 in 1.0.0 and its own license,
 * published as a file, from 2.x on.
 */
function licenseOf(nuspecPath) {
  const xml = readFileSync(nuspecPath, 'utf8')
  const expression = /<license\s+type="expression"\s*>([^<]+)<\/license>/i.exec(xml)
  return expression?.[1]?.trim() ?? null
}
