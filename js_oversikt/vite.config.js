import { defineConfig } from 'vite';
import includeHtml from 'vite-plugin-include-html';
import { resolve } from 'path';

export default defineConfig({
  root: '.',
  plugins: [
    includeHtml() // prosesserer <include src="../partials/..."></include>
  ],
  server: {
    open: '/src/pages/index.html'
  },
  build: {
    outDir: 'dist',
    emptyOutDir: true,
    rollupOptions: {
      input: {
        index:   resolve(__dirname, 'src/pages/index.html'),
        budget:  resolve(__dirname, 'src/pages/budget.html'),
        reports: resolve(__dirname, 'src/pages/reports.html')
      }
    }
  }
});
