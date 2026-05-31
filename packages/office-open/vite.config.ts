import nodePolyfills from "@rolldown/plugin-node-polyfills";
import { defineConfig } from "vite-plus";

export default defineConfig({
  pack: {
    entry: [
      "src/index.ts",
      "src/docx.ts",
      "src/pptx.ts",
      "src/xlsx.ts",
      "src/core.ts",
      "src/xml.ts",
      "src/generate.ts",
      "src/schemas/index.ts",
      "src/ai/index.ts",
      "src/cli.ts",
    ],
    plugins: [nodePolyfills()],
    shims: true,
  },
});
