import js from '@eslint/js'
import tseslint from 'typescript-eslint'

/**
 * The network is reachable only from `src/main/ai/`, and only after the person sets up an AI
 * assistant. `src/main/network-boundary.test.ts` checks the same from the source text.
 */
const NETWORK_MESSAGE = 'Network only in src/main/ai: the rest of the app does not access the internet.'
const NETWORK_PATHS = [
  { name: '@anthropic-ai/sdk', message: NETWORK_MESSAGE },
  { name: 'undici', message: NETWORK_MESSAGE },
  { name: '@agentclientprotocol/sdk', message: NETWORK_MESSAGE },
]
const ELECTRON_NET = {
  selector: "ImportDeclaration[source.value='electron'] > ImportSpecifier[imported.name='net']",
  message: NETWORK_MESSAGE,
}
const NETWORK_PATTERNS = [
  {
    group: ['node:http', 'node:https', 'node:http2', 'node:net', 'node:tls', 'node:dgram'],
    message: NETWORK_MESSAGE,
  },
  { group: ['@agentclientprotocol/sdk/*'], message: NETWORK_MESSAGE },
]
const CHILD_PROCESS_PATTERNS = [
  {
    group: ['node:child_process'],
    message: 'Processes only in the sidecar, the system fonts and the AI assistant.',
  },
]
const NETWORK_GLOBALS = ['fetch', 'WebSocket', 'EventSource', 'XMLHttpRequest'].map((name) => ({
  name,
  message: NETWORK_MESSAGE,
}))

/** The development server, in development only. */
const DEV_SERVER_PROXY = 'src/main/pdf-scheme.ts'
/** Reads the PDF from the viewer's own scheme, inside the isolated frame. */
const PDF_VIEWER_FETCH = 'src/renderer/pdf-viewer/ranged-source.ts'
const PROCESS_SPAWNERS = ['src/main/sidecar/client.ts', 'src/main/system-fonts.ts']

const restrictImports = (paths, patterns) => ({ 'no-restricted-imports': ['error', { paths, patterns }] })

/**
 * Besides the usual rules, this file makes the architecture boundaries checkable by the linter. The
 * golden rule, "services/ does not import electron or react", becomes a build error instead of a
 * recommendation in a document.
 */
export default tseslint.config(
  { ignores: ['out/**', 'dist/**', 'node_modules/**', 'sidecar/**', 'resources/**', '.nix/**', '*.csv'] },

  js.configs.recommended,
  ...tseslint.configs.recommended,

  {
    rules: {
      '@typescript-eslint/consistent-type-imports': ['error', { fixStyle: 'inline-type-imports' }],
      '@typescript-eslint/no-unused-vars': ['error', { argsIgnorePattern: '^_' }],
      eqeqeq: ['error', 'always'],
      'no-console': ['warn', { allow: ['warn', 'error'] }],
    },
  },

  {
    // The pure logic layer: it must run in both main and renderer, and be testable without
    // Electron. Any import from here would break that.
    files: ['src/services/**/*.ts', 'src/shared/**/*.ts'],
    rules: {
      'no-restricted-imports': [
        'error',
        {
          paths: [
            { name: 'electron', message: 'Pure layer: move the Electron usage to src/main.' },
            { name: 'react', message: 'Pure layer: no React here.' },
            { name: 'react-dom', message: 'Pure layer: no React here.' },
            ...NETWORK_PATHS,
          ],
          patterns: [
            { group: ['node:*'], message: 'Pure layer: no Node APIs, this code runs in the renderer.' },
            {
              group: ['@main/*', '@renderer/*'],
              message: 'A pure layer depends on neither main nor renderer.',
            },
            ...NETWORK_PATTERNS,
          ],
        },
      ],
    },
  },

  {
    files: ['src/renderer/**/*.ts', 'src/renderer/**/*.tsx'],
    rules: {
      'no-restricted-imports': [
        'error',
        {
          paths: [
            { name: 'electron', message: 'The renderer talks to main only through window.api.' },
            ...NETWORK_PATHS,
          ],
          patterns: [
            { group: ['node:*'], message: 'The renderer has no Node.js; use window.api.' },
            { group: ['@main/*'], message: 'The renderer does not import from the main process.' },
          ],
        },
      ],
    },
  },

  {
    // The preload is a forwarder: only electron and the shared contracts.
    files: ['src/preload/**/*.ts'],
    rules: {
      'no-restricted-imports': [
        'error',
        {
          paths: NETWORK_PATHS,
          patterns: [
            {
              group: ['@services/*', '@renderer/*', '@main/*'],
              message: 'The preload may only import from @shared; keep it minimal.',
            },
            ...NETWORK_PATTERNS,
          ],
        },
      ],
    },
  },

  {
    // Main has no other diagnostic channel: what it logs on startup is what is left to investigate
    // "opens wrong only on that machine".
    files: ['src/main/**/*.ts'],
    rules: { 'no-console': ['warn', { allow: ['info', 'warn', 'error'] }] },
  },

  {
    files: ['src/main/**/*.ts'],
    ignores: ['src/main/ai/**', DEV_SERVER_PROXY, ...PROCESS_SPAWNERS],
    rules: restrictImports(NETWORK_PATHS, [...NETWORK_PATTERNS, ...CHILD_PROCESS_PATTERNS]),
  },
  { files: PROCESS_SPAWNERS, rules: restrictImports(NETWORK_PATHS, NETWORK_PATTERNS) },
  { files: [DEV_SERVER_PROXY], rules: restrictImports([], CHILD_PROCESS_PATTERNS) },

  {
    files: ['src/**/*.ts', 'src/**/*.tsx'],
    ignores: ['src/main/ai/**', DEV_SERVER_PROXY, PDF_VIEWER_FETCH, '**/*.test.ts', '**/*.test.tsx'],
    rules: {
      'no-restricted-globals': ['error', ...NETWORK_GLOBALS],
      'no-restricted-syntax': ['error', ELECTRON_NET],
    },
  },

  {
    // Function size and complexity. The limits came from measuring the code (p95 complexity 13,
    // lines 44, parameters 4).
    files: ['src/**/*.ts', 'src/**/*.tsx'],
    ignores: ['**/*.test.ts', '**/*.test.tsx'],
    rules: {
      complexity: ['error', 15],
      'max-lines-per-function': ['error', { max: 80, skipBlankLines: true, skipComments: true }],
      'max-depth': ['error', 4],
      'max-params': ['error', 5],
    },
  },

  {
    files: ['**/*.test.ts', '**/*.test.tsx'],
    rules: { 'no-restricted-imports': 'off' },
  },

  {
    // Build scripts: they run on Node directly, outside the bundle, and talk to the user through
    // the terminal; `console` there is the interface, not leftover debugging.
    files: ['scripts/**/*.mjs'],
    languageOptions: {
      globals: { process: 'readonly', console: 'readonly', Buffer: 'readonly' },
    },
    rules: { 'no-console': 'off' },
  },
)
