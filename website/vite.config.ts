import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import path from "path";

export default defineConfig({
  plugins: [react()],
  root: path.resolve(__dirname),
  base: "./",
  resolve: {
    alias: {
      "@": path.resolve(__dirname, "./src")
    }
  },
  css: {
    postcss: path.resolve(__dirname, "./postcss.config.js")
  },
  build: {
    outDir: "dist",
    emptyOutDir: true
  },
  server: {
    port: 5174,
    strictPort: false,
    open: true
  }
});
