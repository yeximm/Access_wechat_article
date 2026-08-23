import { fileURLToPath, URL } from 'node:url'

import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'
import vueDevTools from 'vite-plugin-vue-devtools'

function removeImpeccableLiveScript() {
  return {
    name: 'remove-impeccable-live-script',
    apply: 'build' as const,
    transformIndexHtml(html: string) {
      // 构建静态页面时删除 live 调试脚本，避免桌面壳额外请求 localhost:8400。
      return html.replace(
        /\s*<!-- impeccable-live-start -->[\s\S]*?<!-- impeccable-live-end -->/g,
        '',
      )
    },
  }
}

// https://vite.dev/config/
export default defineConfig({
  base: './',
  server: {
    proxy: {
      // Chrome/Vite 开发模式下把本地 API 转发给 Python dev_server.py。
      '/api': 'http://127.0.0.1:8766',
    },
  },
  plugins: [
    vue(),
    vueDevTools(),
    removeImpeccableLiveScript(),
  ],
  build: {
    outDir: '../src/webview',
    emptyOutDir: true,
  },
  resolve: {
    alias: {
      '@': fileURLToPath(new URL('./src', import.meta.url)),
    },
  },
})
