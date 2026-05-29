import path from "node:path";

import nodePolyfills from "@rolldown/plugin-node-polyfills";
import { defineConfig } from "vite-plus";

const src = path.resolve("src");

export default defineConfig({
  pack: {
    entry: ["src/index.ts"],
    platform: "neutral",
    plugins: [nodePolyfills()],
    shims: true,
  },
  resolve: {
    alias: {
      "@file": path.resolve(src, "file"),
      "@export": path.resolve(src, "export"),
      "@util": path.resolve(src, "util"),
      tests: path.resolve(src, "tests"),
    },
  },
});
