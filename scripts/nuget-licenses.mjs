/**
 * Licenças dos pacotes NuGet resolvidos, lidas do cache local, para o portão
 * (`sidecar.mjs licenses`) e o aviso (`notices.mjs`) não divergirem.
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

/** `[{ name, version, license }]`, ordenado. `license` é `null` quando não declarada. */
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
    // O cache do NuGet usa nomes em minúsculas.
    const directory = join(folder, name.toLowerCase(), version.toLowerCase())
    if (!existsSync(directory)) continue

    const file = readdirSync(directory).find((entry) => entry.toLowerCase().endsWith('.nuspec'))
    if (file !== undefined) return join(directory, file)
  }
  return null
}

/**
 * Só expressão SPDX conta: `SixLabors.Fonts` é Apache-2.0 na 1.0.0 e licença
 * própria, publicada como arquivo, da 2.x em diante.
 */
function licenseOf(nuspecPath) {
  const xml = readFileSync(nuspecPath, 'utf8')
  const expression = /<license\s+type="expression"\s*>([^<]+)<\/license>/i.exec(xml)
  return expression?.[1]?.trim() ?? null
}
