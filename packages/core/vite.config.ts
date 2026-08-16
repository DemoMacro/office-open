import nodePolyfills from "@rolldown/plugin-node-polyfills";
import { defineConfig } from "vite-plus";

export default defineConfig({
  pack: {
    entry: [
      "src/index.ts",
      "src/theme/index.ts",
      "src/smartart/index.ts",
      "src/chart/index.ts",
      "src/drawing/index.ts",
      "src/vector/index.ts",
      "src/descriptor/index.ts",
      "src/patch/index.ts",
      "src/util/index.ts",
    ],
    plugins: [nodePolyfills()],
    shims: true,
  },
  test: {
    benchmark: {
      include: ["src/**/*.bench.ts"],
    },
  },
});
