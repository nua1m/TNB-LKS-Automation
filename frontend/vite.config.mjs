import path from "node:path";

import react from "@vitejs/plugin-react";
import { defineConfig } from "vite";

const frontendRoot = process.cwd();

export default defineConfig({
  plugins: [react()],
  resolve: {
    alias: {
      "@": path.resolve(frontendRoot, "src"),
    },
  },
  base: "./",
  build: {
    outDir: path.resolve(frontendRoot, "..", "web_ui"),
    emptyOutDir: true,
  },
});
