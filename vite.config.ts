import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vite.dev/config/
// Change base to '/STEtruc/' for GitHub Pages deployment under that repo name
export default defineConfig({
  plugins: [react()],
  base: '/STEtruc/',
})
