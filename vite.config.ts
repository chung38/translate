import tailwindcss from '@tailwindcss/vite';
import react from '@vitejs/plugin-react';
import path from 'path';
import {defineConfig, loadEnv} from 'vite';

export default defineConfig(({mode}) => {
  const env = loadEnv(mode, '.', '');
  return {
    plugins: [react(), tailwindcss()],
    define: {
      'process.env.GEMINI_API_KEY': JSON.stringify(env.GEMINI_API_KEY),
    },
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
            'excel-lib': ['xlsx'],
            'vendor': ['react', 'react-dom', 'lucide-react', 'framer-motion'],
          },
        },
      },
      chunkSizeWarningLimit: 2000,
    },
    server: {
      // HMR is disabled in AI Studio via DISABLE_HMR env var.
      // Do not modifyâfile watching is disabled to prevent flickering during agent edits.
      hmr: process.env.DISABLE_HMR !== 'true',
    },
  };
});
