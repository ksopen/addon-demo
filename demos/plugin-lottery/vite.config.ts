import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'
import { ViteFilemanager } from 'filemanager-plugin'

// https://vitejs.dev/config/
export default defineConfig({
  base: './',
  plugins: [
    vue(),
    ViteFilemanager({
      events: {
        end: {
          zip: {
            items: [{
              source: 'dist/*',
              destination: 'plugin.zip',
              type: 'zip'
            }]
          }
        }
      }
    })
  ]
})
