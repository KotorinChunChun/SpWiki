require('@rushstack/eslint-config/patch/modern-module-resolution');
module.exports = {
  extends: ['@microsoft/eslint-config-spfx/lib/profiles/default'],
  parserOptions: { tsconfigRootDir: __dirname },
  overrides: [
    {
      files: ['**/*.ts', '**/*.tsx'],
      parser: '@typescript-eslint/parser',
      parserOptions: {
        project: './tsconfig.json',
        ecmaVersion: 2018,
        sourceType: 'module'
      },
      rules: {
        '@rushstack/no-new-null': 1,
        '@rushstack/security/no-unsafe-regexp': 1,
        '@typescript-eslint/explicit-function-return-type': [
          1,
          {
            allowExpressions: true,
            allowTypedFunctionExpressions: true,
            allowHigherOrderFunctions: false
          }
        ],
        '@typescript-eslint/no-explicit-any': 1,
        '@typescript-eslint/no-floating-promises': 2,
        '@typescript-eslint/no-unused-vars': [1, { vars: 'all', args: 'none' }],
        '@typescript-eslint/no-inferrable-types': 0,
        '@typescript-eslint/no-empty-interface': 0,
        'eqeqeq': 1,
        'no-var': 2,
        'prefer-const': 1,
        'strict': [2, 'never'],
        '@microsoft/spfx/no-require-ensure': 2
      }
    }
  ]
};
