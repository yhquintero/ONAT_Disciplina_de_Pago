import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import basicSsl from '@vitejs/plugin-basic-ssl';

const base = process.env.VITE_PORTABLE_BUILD === 'true'
  ? './'
  : process.env.VITE_BASE_PATH || '/';
const useLocalHttps = process.env.LOCAL_HTTPS === 'true';

export default defineConfig({
  base,
  plugins: [react(), ...(useLocalHttps ? [basicSsl()] : [])],
  assetsInclude: ['**/*.xlsx', '**/*.xls', '**/*.xlsm', '**/*.xlsb'],
  // ExcelJS is loaded only when the user exports to XLSX; its ~256 kB gzip bundle is intentional.
  build: { chunkSizeWarningLimit: 1000 },
  server: {
    host: '0.0.0.0',
    allowedHosts: true,
  },
  preview: {
    host: '0.0.0.0',
    allowedHosts: true,
  },
});
