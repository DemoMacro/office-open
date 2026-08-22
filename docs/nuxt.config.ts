import { readdirSync, readFileSync, statSync } from "node:fs";
import { join } from "node:path";

import { useNuxt } from "@nuxt/kit";

// clientBundle.scan only sees icons referenced as literals inside vite modules;
// icons from docus .navigation.yml and content frontmatter live in the content
// dump instead, and with provider "server" those fall back to
// /api/_nuxt_icon/* at runtime — which does not exist in a fully static deploy.
// Collect every icon literal ourselves and hand the list to the client bundle.
function collectIconNames(): string[] {
  const names = new Set<string>();
  const walk = (dir: string) => {
    for (const entry of readdirSync(dir)) {
      const full = join(dir, entry);
      if (statSync(full).isDirectory()) {
        if (entry !== "node_modules") walk(full);
      } else if (/\.(vue|ts|md|yml|json)$/.test(entry)) {
        for (const match of readFileSync(full, "utf8").matchAll(
          /\bi-[a-z][a-z0-9]*(?:-[a-z0-9]+)+/g,
        )) {
          names.add(match[0]);
        }
      }
    }
  };
  for (const dir of ["app", "content"]) {
    try {
      walk(join(__dirname, dir));
    } catch {
      // directory missing in this workspace layout — nothing to scan
    }
  }
  // clientBundle.icons expects "collection:name"; component usage is "i-collection-name".
  const collections = ["vscode-icons", "simple-icons", "lucide", "custom"];
  return [...names]
    .map((icon) => {
      const body = icon.slice(2);
      for (const collection of collections) {
        if (body.startsWith(collection + "-")) {
          return `${collection}:${body.slice(collection.length + 1)}`;
        }
      }
      return null;
    })
    .filter((name): name is string => name !== null);
}

export default defineNuxtConfig({
  extends: ["docus"],
  modules: ["@nuxtjs/i18n"],
  css: ["~/assets/css/main.css"],

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

  // unifont initializes its google/googleicons providers on every build,
  // fetching fonts.google.com metadata (with retry backoff) even though this
  // site requests no Google web font or material-symbols icon. Disable both;
  // icon rendering is fully covered by @nuxt/icon with local collections.
  fonts: {
    providers: {
      google: false,
      googleicons: false,
    },
  },

  // Resolve icons from locally installed @iconify-json/* collections instead
  // of the iconify CDN (docus defaults provider to "iconify", which makes
  // every visitor's browser call api.iconify.design at runtime).
  icon: {
    provider: "server",
    serverBundle: "local",
    clientBundle: {
      icons: collectIconNames(),
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
    // docus pulls in @nuxt/image but neither this site nor the docus theme
    // renders a single NuxtImg/NuxtPicture; dropping the module keeps ipx and
    // the sharp binary wiring out of the server bundle.
    "modules:before"() {
      const { modules } = useNuxt().options;
      const image = modules.indexOf("@nuxt/image");
      if (image !== -1) modules.splice(image, 1);
    },
    "nitro:config"(nitroConfig) {
      nitroConfig.handlers = nitroConfig.handlers?.filter((h) => h?.route !== "/api/search");
    },
  },
});
