import path from "node:path";

import nodePolyfills from "@rolldown/plugin-node-polyfills";
import { defineConfig, type PluginOption } from "vite-plus";

const src = path.resolve("src");

export default defineConfig({
  pack: {
    entry: ["src/index.ts"],
    plugins: [nodePolyfills()] as PluginOption[],
    shims: true,
  },
  resolve: {
    alias: {
      "@parts": path.resolve(src, "parts"),
      "@shared": path.resolve(src, "shared"),
      "@util": path.resolve(src, "util"),
      tests: path.resolve(src, "tests"),
    },
  },
});
