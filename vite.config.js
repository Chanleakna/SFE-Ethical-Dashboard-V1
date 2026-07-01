import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import { viteSingleFile } from 'vite-plugin-singlefile'

// Single-file build: inlines all JS/CSS into one dist/index.html so it can be
// pasted into the Apps Script `index` HTML file (Apps Script can't serve separate
// asset files). Keeps the xlsx CDN <script> external (loaded at runtime).
export default defineConfig({
  plugins: [react(), viteSingleFile()],
})
