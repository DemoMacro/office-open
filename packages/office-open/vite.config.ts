import nodePolyfills from "@rolldown/plugin-node-polyfills";
import { defineConfig, type PluginOption } from "vite-plus";

export default defineConfig({
  pack: {
    entry: [
      "src/index.ts",
      "src/docx.ts",
      "src/document.ts",
      "src/pptx.ts",
      "src/presentation.ts",
      "src/xlsx.ts",
      "src/workbook.ts",
      "src/doc.ts",
      "src/xls.ts",
      "src/ppt.ts",
      "src/rtf.ts",
      "src/formats.ts",
      "src/detect.ts",
      "src/parse.ts",
      "src/core.ts",
      "src/xml.ts",
      "src/generate.ts",
      "src/schemas/index.ts",
      "src/ai/index.ts",
      "src/convert/index.ts",
      "src/cli.ts",
    ],
    plugins: [nodePolyfills()] as PluginOption[],
    shims: true,
  },
});
