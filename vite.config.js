import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  optimizeDeps: {
    include: ['exceljs', 'jszip']
  },
  resolve: {
    alias: {
      './zlib_bindings': './zlib_bindings'
    }
  },
  define: {
    global: 'globalThis'
  }
})
