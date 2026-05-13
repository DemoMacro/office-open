export default defineNuxtConfig({
    extends: ["docus"],
    modules: ["@nuxtjs/i18n"],

    vite: {
        optimizeDeps: {
            include: ["@vue/devtools-core", "@vue/devtools-kit", "remark-emoji", "remark-mdc"],
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
        domain: "https://office-open.com",
        title: "Office Open",
        description:
            "Generate Office Open XML documents (.docx, .pptx, .xlsx) with JavaScript/TypeScript",
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
    },

    hooks: {
        "nitro:config"(nitroConfig) {
            nitroConfig.handlers = nitroConfig.handlers?.filter((h) => h?.route !== "/api/search");
        },
    },
});
