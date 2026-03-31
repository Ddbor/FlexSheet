import path from "node:path";
import { fileURLToPath } from "node:url";
import { defineConfig } from "vite";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

const alias = {
  flexsheet: path.resolve(__dirname, "packages/flexsheet/src/index.ts"),
  "@flexsheet/core": path.resolve(__dirname, "packages/core/src/index.ts"),
  "@flexsheet/shared": path.resolve(__dirname, "packages/shared/src/index.ts"),
  "@flexsheet/theme": path.resolve(__dirname, "packages/theme/src/index.ts"),
  "@flexsheet/formula": path.resolve(__dirname, "packages/formula/src/index.ts"),
  "@flexsheet/renderer": path.resolve(__dirname, "packages/renderer/src/index.ts"),
  "@flexsheet/selection": path.resolve(__dirname, "packages/selection/src/index.ts"),
  "@flexsheet/editor": path.resolve(__dirname, "packages/editor/src/index.ts"),
  "@flexsheet/scroll": path.resolve(__dirname, "packages/scroll/src/index.ts"),
  "@flexsheet/import-export": path.resolve(__dirname, "packages/import-export/src/index.ts"),
  "@flexsheet/toolbar": path.resolve(__dirname, "packages/toolbar/src/index.ts"),
} as const;

export default defineConfig({
  root: ".",
  resolve: {
    alias,
  },
  server: {
    port: 5173,
    open: true,
  },
  build: {
    lib: {
      entry: path.resolve(__dirname, "packages/flexsheet/src/index.ts"),
      name: "FlexSheet",
      formats: ["es", "cjs", "iife"],
      fileName: (format): string => {
        if (format === "es") return "flexsheet.es.js";
        if (format === "cjs") return "flexsheet.cjs";
        return "flexsheet.iife.js";
      },
    },
    rollupOptions: {
      output: {
        extend: false,
        exports: "named",
      },
    },
    emptyOutDir: true,
  },
});
