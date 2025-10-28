// Timestamp: 2025-10-28, Version: 2.0.1.0
import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import path from 'path';

export default defineConfig({
  // FIX: Set the base path to correctly resolve assets when deployed to the 
  // GitHub Pages subfolder (https://vsblanco.github.io/Student-Retention-Add-in/react/).
  base: '/Student-Retention-Add-in/react/', 
  plugins: [react()],
  css: {
    // Note: path.resolve(__dirname, 'postcss.config.cjs') might not work 
    // depending on the execution environment; './postcss.config.cjs' might be safer.
    postcss: path.resolve(__dirname, 'postcss.config.cjs'),
  },
});