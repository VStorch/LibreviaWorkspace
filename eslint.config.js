import js from '@eslint/js'
import tseslint from 'typescript-eslint'

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
            { name: 'electron', message: 'Camada pura: mova o uso de Electron para src/main.' },
            { name: 'react', message: 'Camada pura: sem React aqui.' },
            { name: 'react-dom', message: 'Camada pura: sem React aqui.' },
          ],
          patterns: [
            { group: ['node:*'], message: 'Camada pura: sem APIs do Node — este código roda no renderer.' },
            {
              group: ['@main/*', '@renderer/*'],
              message: 'Camada pura não depende de main nem de renderer.',
            },
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
          paths: [{ name: 'electron', message: 'O renderer fala com o main apenas por window.api.' }],
          patterns: [
            { group: ['node:*'], message: 'O renderer não tem Node.js — use window.api.' },
            { group: ['@main/*'], message: 'O renderer não importa do processo main.' },
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
          patterns: [
            {
              group: ['@services/*', '@renderer/*', '@main/*'],
              message: 'O preload só pode importar de @shared — mantenha-o mínimo.',
            },
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
