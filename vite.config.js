import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [react()],
  base: './', // 设置基础路径为相对路径
  server: {
    host: '0.0.0.0',
    port: 3000,
    open: true,
  }
})
