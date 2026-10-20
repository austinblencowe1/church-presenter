import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
// @ts-expect-error type error without @types/node package
import process from "node:process";
// @ts-expect-error type error without @types/node package
import { readFileSync } from "node:fs";
// @ts-expect-error type error without @types/node package
import { resolve } from "node:path";
const host = process.env.TAURI_DEV_HOST;

// https://vite.dev/config/
export default defineConfig(() => ({
  plugins: [
    react(),
    {
      name: "presenter-runtime-assets",
      generateBundle() {
        for (const fileName of ["cloud-section.js", "gdrive-client-id.txt"]) {
          this.emitFile({
            type: "asset",
            fileName,
            source: readFileSync(resolve(process.cwd(), fileName)),
          });
        }
      },
    },
  ],
  build: {
    // The legacy presenter remains the app's source of truth. Build the
    // audience window alongside it so Tauri can open it in production too.
    rollupOptions: {
      input: {
        main: resolve(process.cwd(), "index.html"),
        audience: resolve(process.cwd(), "audience.html"),
      },
    },
  },

  // Vite options tailored for Tauri development and only applied in `tauri dev` or `tauri build`
  //
  // 1. prevent Vite from obscuring rust errors
  clearScreen: false,
  // 2. tauri expects a fixed port, fail if that port is not available
  server: {
    port: 1420,
    strictPort: true,
    host: host || false,
    hmr: host
      ? {
          protocol: "ws",
          host,
          port: 1421,
        }
      : undefined,
    watch: {
      // 3. tell Vite to ignore watching `src-tauri`
      ignored: ["**/src-tauri/**"],
    },
  },
}));
