import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// base 設成 repo 名稱，GitHub Pages 才能正確載入資源
export default defineConfig({
  plugins: [react()],
  base: '/heheyen-recon/',
})
