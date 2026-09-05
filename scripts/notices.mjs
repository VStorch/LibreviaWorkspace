/**
 * Gera `THIRD-PARTY-NOTICES.md` — o que vai junto com o instalador.
 *
 * Cobre os **três** conjuntos que de fato são distribuídos, e não só o primeiro:
 *
 *  1. dependências npm de produção;
 *  2. o próprio Electron, que é `devDependency` no `package.json` mas viaja
 *     inteiro dentro do instalador — junto com Chromium e Node.js;
 *  3. os pacotes NuGet ligados ao sidecar .NET, que é publicado self-contained e
 *     por isso leva o runtime junto.
 *
 * O portão `licenses:check` olha só o conjunto 1. Deixar os outros dois de fora
 * do arquivo de avisos daria a impressão de conformidade sem a conformidade.
 */

import { execFileSync } from 'node:child_process'
import { readFileSync, writeFileSync } from 'node:fs'
import { createRequire } from 'node:module'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { collectNuGetLicenses, nugetAssetsExist } from './nuget-licenses.mjs'

const root = join(dirname(fileURLToPath(import.meta.url)), '..')
const require = createRequire(import.meta.url)

function npmPackages() {
  // O próprio JS do license-checker, rodado por este Node — e não `npx`. No
  // Windows `npx` é `npx.cmd`, e `execFile` sem shell não executa um `.cmd`:
  // era o `spawnSync npx ENOENT` que derrubava o instalador lá. Resolver pelo
  // `require` acha o pacote onde quer que o npm o tenha içado, e dispensa
  // shell — que traria de volta o problema de caminho com espaço.
  const checker = require.resolve('license-checker-rseidelsohn/bin/license-checker-rseidelsohn.js')

  const json = execFileSync(
    process.execPath,
    [checker, '--production', '--excludePrivatePackages', '--json', '--start', root],
    { encoding: 'utf8', maxBuffer: 32 * 1024 * 1024 },
  )

  return Object.entries(JSON.parse(json))
    .map(([key, value]) => {
      const at = key.lastIndexOf('@')
      return {
        name: key.slice(0, at),
        version: key.slice(at + 1),
        license: value.licenses ?? '(não declarada)',
        repository: value.repository ?? null,
      }
    })
    .sort((a, b) => a.name.localeCompare(b.name))
}

function electronRuntime() {
  const manifest = JSON.parse(readFileSync(join(root, 'node_modules/electron/package.json'), 'utf8'))
  return manifest.version
}

/**
 * As fontes empacotadas.
 *
 * Escritas à mão, e não lidas de `resources/fonts/`: o que interessa ao aviso é
 * a **família** e quem a assina, não os vinte arquivos de variante. A conferência
 * de que os arquivos existem é do teste `src/services/document/fonts.test.ts`.
 */
function bundledFonts() {
  return [
    ['Carlito (substitui Calibri)', 'OFL-1.1', 'https://github.com/googlefonts/carlito'],
    ['Caladea (substitui Cambria)', 'OFL-1.1', 'https://github.com/huertatipografica/Caladea'],
    ['Liberation Sans (substitui Arial)', 'OFL-1.1', 'https://github.com/liberationfonts'],
    ['Liberation Serif (substitui Times New Roman)', 'OFL-1.1', 'https://github.com/liberationfonts'],
    ['Liberation Mono (substitui Courier New)', 'OFL-1.1', 'https://github.com/liberationfonts'],
  ]
}

function table(rows, columns) {
  const header = `| ${columns.join(' | ')} |`
  const divider = `| ${columns.map(() => '---').join(' | ')} |`
  return [header, divider, ...rows.map((cells) => `| ${cells.join(' | ')} |`)].join('\n')
}

function build() {
  const npm = npmPackages()
  const nuget = nugetAssetsExist() ? collectNuGetLicenses() : []

  const sections = [
    '# Avisos de terceiros',
    '',
    'O Librevia é distribuído com os componentes abaixo, cada um sob a sua própria',
    'licença. Este arquivo é gerado por `npm run notices` e não deve ser editado à mão.',
    '',
    `Gerado em ${new Date().toISOString().slice(0, 10)}.`,
    '',
    '## Plataforma',
    '',
    'O instalador leva o Electron inteiro, e com ele o Chromium e o Node.js.',
    '',
    table(
      [
        [`Electron ${electronRuntime()}`, 'MIT', 'https://github.com/electron/electron'],
        [
          'Chromium (via Electron)',
          'BSD-3-Clause e outras',
          'https://chromium.googlesource.com/chromium/src',
        ],
        ['Node.js (via Electron)', 'MIT', 'https://github.com/nodejs/node'],
      ],
      ['Componente', 'Licença', 'Origem'],
    ),
    '',
    `## Dependências npm de produção (${npm.length})`,
    '',
    table(
      npm.map(({ name, version, license, repository }) => [`${name} ${version}`, license, repository ?? '—']),
      ['Pacote', 'Licença', 'Origem'],
    ),
    '',
    '## Fontes',
    '',
    'O instalador leva cinco famílias metricamente compatíveis com as fontes que os',
    'documentos pedem. Sem elas o sistema substitui por conta própria, a métrica',
    'muda e o documento pagina diferente — ver `src/services/document/fonts.ts`.',
    '',
    table(bundledFonts(), ['Família', 'Licença', 'Origem']),
    '',
  ]

  if (nuget.length > 0) {
    sections.push(
      `## Pacotes NuGet do serviço de formatos (${nuget.length})`,
      '',
      'O serviço `.NET` é publicado self-contained: o runtime do .NET viaja dentro do',
      'instalador, sob a licença MIT da Microsoft.',
      '',
      table(
        [
          ['Runtime do .NET 10', 'MIT', 'https://github.com/dotnet/runtime'],
          ...nuget.map(({ name, version, license }) => [
            `${name} ${version}`,
            license ?? '(não declarada)',
            `https://www.nuget.org/packages/${name}`,
          ]),
        ],
        ['Componente', 'Licença', 'Origem'],
      ),
      '',
    )
  } else {
    sections.push(
      '## Pacotes NuGet do serviço de formatos',
      '',
      '_Não listados: `project.assets.json` ausente. Rode `npm run sidecar:build` antes._',
      '',
    )
  }

  const target = join(root, 'THIRD-PARTY-NOTICES.md')
  writeFileSync(target, sections.join('\n'))
  console.log(`avisos gravados em ${target} — ${npm.length} pacotes npm, ${nuget.length} NuGet`)
}

build()
