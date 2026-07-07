export default defineNuxtConfig({
  extends: ["docus"],
  modules: ["@nuxtjs/i18n"],

  // Nitro's server-bundle esbuild plugin defaults to target "es2019"
  // (rollup/index.mjs), which rejects top-level await emitted by some
  // dependencies bundled under the Vercel preset. Raise it to "es2022"
  // (TLA is part of ES2022) so the prerenderer stops failing on Vercel.
  nitro: {
    esbuild: {
      options: {
        target: "es2022",
      },
    },
  },

  vite: {
    optimizeDeps: {
      include: [
        "@ai-sdk/vue",
        "@noble/hashes/legacy.js",
        "@noble/hashes/sha2.js",
        "@noble/hashes/utils.js",
        "@office-open/docx > @office-open/core",
        "@office-open/pptx > @office-open/core",
        "@office-open/xlsx > @office-open/core",
        "@vue/devtools-core",
        "@vue/devtools-kit",
        "@vueuse/core",
        "ai",
        "fflate",
        "remark-emoji",
        "remark-mdc",
      ],
    },
  },

  i18n: {
    locales: [
      { code: "en", name: "English", file: "en.json" },
      { code: "zh", name: "中文", file: "zh.json" },
    ],
    defaultLocale: "en",
    langDir: "locales",
  },

  llms: {
    domain: "https://www.office-open.com",
    title: "Office Open",
    description:
      "TypeScript toolkit for Office documents — generate, parse, and patch .docx, .pptx, .xlsx with spec-compliant OOXML output",
    full: {
      title: "Office Open — Full Documentation",
      description:
        "Complete API reference and guides for @office-open/docx, @office-open/pptx, and @office-open/core.",
    },
    contentRawMarkdown: {
      excludeCollections: ["landing_en", "landing_zh"],
      rewriteLLMSTxt: true,
    },
  },

  docus: {
    assistant: {
      mcpServer: "/mcp",
      apiPath: "/api/search",
    },
    skills: {
      dir: "skills",
    },
  },

  hooks: {
    "nitro:config"(nitroConfig) {
      nitroConfig.handlers = nitroConfig.handlers?.filter((h) => h?.route !== "/api/search");
    },
  },
});
