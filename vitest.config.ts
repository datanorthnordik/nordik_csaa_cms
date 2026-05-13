import react from '@vitejs/plugin-react'
import { defineConfig } from 'vitest/config'

export default defineConfig({
  plugins: [react()],
  test: {
    globals: true,
    include: ['src/**/*.{test,spec}.{ts,tsx}'],
    exclude: ['coverage/**', 'cypress/**', 'dist/**'],
    environment: 'jsdom',
    setupFiles: [],
    deps: {
      optimizer: {
        web: {
          include: [
            '@emotion/react',
            '@emotion/styled',
            '@mui/icons-material/MoreVert',
            '@mui/material/Dialog',
            '@mui/material/DialogActions',
            '@mui/material/DialogContent',
            '@mui/material/DialogContentText',
            '@mui/material/DialogTitle',
            '@mui/material/IconButton',
            '@mui/material/Menu',
            '@mui/material/MenuItem',
            'react-dom',
            'react-hot-toast',
            'react-i18next',
            'react-redux',
            'react-router-dom',
          ],
        },
      },
    },
    coverage: {
      provider: 'v8',
      reporter: ['text-summary', 'json-summary'],
      exclude: [
        'coverage/',
        'cypress/',
        'dist/',
        'node_modules/',
        'src/assets/**',
        'src/i18n/**',
        '**/*.config.*',
        '**/*.d.ts',
        '**/*.test.*',
      ],
    },
  },
})
