import path from "node:path";

import nodePolyfills from "@rolldown/plugin-node-polyfills";
import { defineConfig } from "vite-plus";

const src = path.resolve("src");

export default defineConfig({
  pack: {
    entry: ["src/index.ts", "src/generate.ts", "src/parse.ts", "src/patch.ts"],
    plugins: [nodePolyfills()],
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
