// Flat config (ESLint 9 / typescript-eslint 8).
// Pragmatic baseline: JS + TypeScript "recommended" without type-aware rules,
// so the existing codebase passes without a full type-checked lint pass.
import js from '@eslint/js';
import tseslint from 'typescript-eslint';
import globals from 'globals';

export default tseslint.config(
  {
    ignores: ['dist/**', 'node_modules/**', 'coverage/**'],
  },
  js.configs.recommended,
  ...tseslint.configs.recommended,
  {
    files: ['**/*.ts'],
    languageOptions: {
      ecmaVersion: 2022,
      sourceType: 'module',
      globals: {
        ...globals.node,
      },
    },
    rules: {
      // Server logging goes through the logger; keep stray console calls out of
      // request paths. Intentional bootstrap/shutdown logging is opted in with
      // an inline eslint-disable.
      'no-console': 'error',
      // Allow intentionally-unused args/vars prefixed with underscore.
      '@typescript-eslint/no-unused-vars': [
        'warn',
        {
          argsIgnorePattern: '^_',
          varsIgnorePattern: '^_',
          caughtErrorsIgnorePattern: '^_',
        },
      ],
    },
  },
  {
    files: ['test/**/*.ts'],
    languageOptions: {
      globals: {
        ...globals.node,
      },
    },
  },
  {
    // These serializer modules are a parity port of the established frontend
    // implementation. Keep their current compiler boundary explicit while the
    // backend types are tightened incrementally; runtime behavior is covered by
    // the cross-format fidelity suites.
    files: ['src/export/serialization/**/*.ts'],
    rules: {
      '@typescript-eslint/ban-ts-comment': 'off',
      'no-control-regex': 'off',
      'no-misleading-character-class': 'off',
    },
  },
);
