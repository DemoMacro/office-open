import path from "node:path";

import nodePolyfills from "@rolldown/plugin-node-polyfills";
import { defineConfig } from "vite-plus";

const src = path.resolve("src");

export default defineConfig({
  pack: {
    entry: ["src/index.ts"],
    plugins: [nodePolyfills()],
    shims: true,
  },
  resolve: {
    alias: {
      "@parts": path.resolve(src, "parts"),
      "@shared": path.resolve(src, "shared"),
      "@export": path.resolve(src, "export"),
      "@util": path.resolve(src, "util"),
    },
  },
});
