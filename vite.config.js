import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// De lokale proef (npm run lokaal) zet ANHL_LIVE_PROXY aan: dan gaan /api-aanroepen
// naar de live site. Op de bouw heeft dit geen invloed.
const liveProxy = process.env.ANHL_LIVE_PROXY === '1'
  ? { '/api': { target: 'https://autonomie-nhlstenden.netlify.app', changeOrigin: true, secure: true } }
  : undefined

export default defineConfig({
  plugins: [react()],
  server: { proxy: liveProxy },
  build: {
    outDir: 'dist',
    sourcemap: false
  }
})
