import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  server: {
    fs: {
      strict: true,
      allow: [process.cwd()]
    },
    watch: {
      ignored: [
        '**/node_modules/**',
        '**/.git/**',
        '**/AppData/**',
        '**/Slack/**',
        '**/Cookies*'
      ]
    }
  }
})
