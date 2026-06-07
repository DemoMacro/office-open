import nodePolyfills from "@rolldown/plugin-node-polyfills";
import { defineConfig } from "vite-plus";

export default defineConfig({
  pack: {
    entry: [
      "src/index.ts",
      "src/values.ts",
      "src/theme/index.ts",
      "src/smartart/index.ts",
      "src/chart/index.ts",
      "src/drawingml/index.ts",
      "src/descriptor/index.ts",
      "src/patch/index.ts",
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
