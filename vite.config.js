import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'
import rollupNodePolyFill from 'rollup-plugin-node-polyfills'

export default defineConfig({
  plugins: [vue()],
  resolve: {
    alias: {
      crypto: 'crypto-browserify',
      stream: 'stream-browserify',
      buffer: 'buffer',
    },
  },
  build: {
    rollupOptions: {
      plugins: [rollupNodePolyFill()]
    },
  },
  optimizeDeps: {
    include: ['crypto-browserify', 'stream-browserify', 'buffer'],
  }
})
