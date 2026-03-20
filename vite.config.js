import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  // Reemplazá 'ri-cashflow' con el nombre exacto de tu repositorio en GitHub
  base: '/ri-cashflow/',
})
