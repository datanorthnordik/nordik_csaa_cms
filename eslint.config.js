import js from '@eslint/js'
import globals from 'globals'
import reactHooks from 'eslint-plugin-react-hooks'
import reactRefresh from 'eslint-plugin-react-refresh'
import tseslint from 'typescript-eslint'
import { defineConfig, globalIgnores } from 'eslint/config'

export default defineConfig([
  globalIgnores(['dist', 'coverage']),
  {
    files: ['**/*.{ts,tsx}'],
    extends: [
      js.configs.recommended,
      tseslint.configs.recommended,
      reactHooks.configs.flat.recommended,
      reactRefresh.configs.vite,
    ],
    languageOptions: {
      globals: globals.browser,
    },
    rules: {
      // setState inside useEffect is an intentional pattern used project-wide
      // for syncing form state from loaded data.
      'react-hooks/set-state-in-effect': 'off',
      // Allow _-prefixed variables/args to be intentionally unused.
      '@typescript-eslint/no-unused-vars': [
        'error',
        {
          varsIgnorePattern: '^_',
          argsIgnorePattern: '^_',
          caughtErrorsIgnorePattern: '^_',
          destructuredArrayIgnorePattern: '^_',
        },
      ],
      // Utility functions exported alongside components are intentional in this
      // codebase — downgrade from error to warn so Fast Refresh still flags
      // genuinely mixed files without blocking the build.
      'react-refresh/only-export-components': ['warn', { allowConstantExport: true }],
    },
  },
  // Test files: allow `any` types — needed for mocking and spy setups.
  {
    files: ['**/*.test.ts', '**/*.test.tsx'],
    rules: {
      '@typescript-eslint/no-explicit-any': 'off',
    },
  },
])
