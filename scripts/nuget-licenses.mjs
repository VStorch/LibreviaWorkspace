/**
 * Licenças dos pacotes NuGet resolvidos, lidas do cache local.
 *
 * Mora separado porque tem dois consumidores que não podem divergir: o portão
 * que reprova o build (`sidecar.mjs licenses`) e o arquivo de avisos que vai
 * junto com o instalador (`notices.mjs`). Se cada um lesse por conta própria, um
 * dia o portão aprovaria um pacote que o aviso não menciona.
 */

import { existsSync, readFileSync, readdirSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = join(dirname(fileURLToPath(import.meta.url)), '..')
const assetsPath = join(root, 'sidecar', 'src', 'Librevia.Format', 'obj', 'project.assets.json')

/** Mesma allowlist do lado npm — ver §4.4 do plano. */
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
 * Só expressão SPDX conta.
 *
 * Licença publicada como arquivo em vez de expressão é tratada como não
 * declarada — e isso não é rigor gratuito: `SixLabors.Fonts` é Apache-2.0 na
 * 1.0.0 e licença própria da 2.x em diante, publicada como arquivo. Aceitar
 * "não declarou" como "deve estar tudo bem" abriria o buraco que o portão
 * existe para fechar.
 */
function licenseOf(nuspecPath) {
  const xml = readFileSync(nuspecPath, 'utf8')
  const expression = /<license\s+type="expression"\s*>([^<]+)<\/license>/i.exec(xml)
  return expression?.[1]?.trim() ?? null
}
