import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  base: '/dara-timetable/',
  build: {
    rollupOptions: {
      external: [
        /^https:\/\/www\.gstatic\.com\/firebasejs\/.*/
      ]
    }
  }
})