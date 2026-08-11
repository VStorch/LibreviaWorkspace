import js from '@eslint/js'
import tseslint from 'typescript-eslint'

/**
 * Além das regras usuais, este arquivo torna as fronteiras da arquitetura
 * verificáveis pelo linter. A regra de ouro do plano — "services/ não importa
 * electron nem react" — vira erro de build em vez de recomendação em documento.
 */
export default tseslint.config(
  { ignores: ['out/**', 'dist/**', 'node_modules/**', '*.csv'] },

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
    // Camada de lógica pura: precisa rodar tanto no main quanto no renderer,
    // e ser testável sem Electron. Qualquer import daqui a quebraria.
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
    // O preload é encaminhador: só electron e os contratos compartilhados.
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
    files: ['**/*.test.ts', '**/*.test.tsx'],
    rules: { 'no-restricted-imports': 'off' },
  },
)
