import { readdirSync, readFileSync } from 'node:fs'
import { join, resolve } from 'node:path'
import react from '@vitejs/plugin-react'
import { defineConfig, externalizeDepsPlugin } from 'electron-vite'
import type { Plugin } from 'vite'

const alias = {
  '@shared': resolve('src/shared'),
  '@services': resolve('src/services'),
  '@main': resolve('src/main'),
  '@renderer': resolve('src/renderer'),
}

/**
 * What pdf.js fetches at run time, copied to `out/renderer/pdfjs`: the installer carries only
 * `out/`. The scripting engine (`quickjs`) stays out, since the viewer never runs a PDF's
 * JavaScript, and so do the decoders' plain-JS fallbacks, which WebAssembly makes unused. Its
 * Liberation Sans is version 1, under the GPL: main serves the app's own, under the OFL, instead.
 */
const PDFJS_RESOURCES: Readonly<Record<string, (name: string) => boolean>> = {
  cmaps: () => true,
  standard_fonts: (name) => !name.includes('LIBERATION') && !name.startsWith('LiberationSans'),
  iccs: () => true,
  wasm: (name) => name.startsWith('LICENSE_') || (name.endsWith('.wasm') && !name.startsWith('quickjs')),
}

function pdfjsResources(): Plugin {
  const root = resolve('node_modules/pdfjs-dist')
  return {
    name: 'librevia:pdfjs-resources',
    apply: 'build',
    generateBundle() {
      for (const [folder, wanted] of Object.entries(PDFJS_RESOURCES)) {
        for (const name of readdirSync(join(root, folder)).filter(wanted)) {
          this.emitFile({
            type: 'asset',
            fileName: `pdfjs/${folder}/${name}`,
            source: readFileSync(join(root, folder, name)),
          })
        }
      }
    },
  }
}

export default defineConfig({
  main: {
    resolve: { alias },
    plugins: [externalizeDepsPlugin()],
    build: {
      rollupOptions: { input: { index: resolve('src/main/index.ts') } },
    },
  },
  preload: {
    resolve: { alias },
    plugins: [externalizeDepsPlugin()],
    build: {
      // `sandbox: true` requires a CommonJS preload: sandboxed preloads do not support ESM. Hence
      // the .cjs output here, not .mjs.
      rollupOptions: {
        input: { index: resolve('src/preload/index.ts') },
        output: { format: 'cjs', entryFileNames: '[name].cjs' },
      },
    },
  },
  renderer: {
    root: resolve('src/renderer'),
    resolve: { alias },
    plugins: [react(), pdfjsResources()],
    build: {
      rollupOptions: {
        input: {
          index: resolve('src/renderer/index.html'),
          pdfViewer: resolve('src/renderer/pdf-viewer/index.html'),
        },
      },
    },
  },
})
