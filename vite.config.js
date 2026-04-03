import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  base: '/dara-timetable/',   // ← เพิ่มบรรทัดนี้
  plugins: [react()],
  server: {
    host: '0.0.0.0',
    port: 3000,
  }
})