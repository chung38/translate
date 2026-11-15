import tailwindcss from '@tailwindcss/vite';
import react from '@vitejs/plugin-react';
import path from 'path';
import {defineConfig} from 'vite';

export default defineConfig({
  plugins: [react(), tailwindcss()],
  resolve: {
    alias: {
      '@': path.resolve(__dirname, '.'),
    },
  },
  build: {
    rollupOptions: {
      // 兩個頁面：/ 是「移次搞定」外殼（index.html），/translate/ 是文件翻譯（React）
      input: {
        main: path.resolve(__dirname, 'index.html'),
        translate: path.resolve(__dirname, 'translate/index.html'),
      },
      output: {
        manualChunks: {
          'pdf-lib': ['pdfjs-dist', 'jspdf'],
          'vendor': ['react', 'react-dom', 'lucide-react', 'framer-motion'],
        },
      },
    },
    chunkSizeWarningLimit: 2000,
  },
  server: {
    // npm run dev 只跑畫面；翻譯 API 在 functions/，要另開 npm run preview（port 8788）才有
    proxy: {
      '/api': 'http://localhost:8788',
    },
  },
});
