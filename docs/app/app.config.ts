export default defineAppConfig({
  header: {
    title: "Office Open",
  },
  seo: {
    titleTemplate: "%s - Office Open",
    title: "Office Open",
    description:
      "TypeScript toolkit for Office documents — generate, parse, and patch .docx, .pptx, .xlsx with spec-compliant OOXML output",
  },
  github: {
    url: "https://github.com/DemoMacro/office-open",
    branch: "main",
    rootDir: "docs",
  },
  navigation: { sub: "header" },
  ui: {
    prose: {
      codeIcon: {
        // Office document types
        docx: "i-vscode-icons-file-type-word",
        pptx: "i-vscode-icons-file-type-powerpoint",
        xlsx: "i-vscode-icons-file-type-excel",
        // Markup & data
        json: "i-vscode-icons-file-type-json",
        jsonc: "i-vscode-icons-file-type-json",
        html: "i-vscode-icons-file-type-html",
        htm: "i-vscode-icons-file-type-html",
        xml: "i-vscode-icons-file-type-xml",
        markdown: "i-vscode-icons-file-type-markdown",
        md: "i-vscode-icons-file-type-markdown",
        mdc: "i-vscode-icons-file-type-markdown",
        css: "i-vscode-icons-file-type-css",
        scss: "i-vscode-icons-file-type-css",
        less: "i-vscode-icons-file-type-css",
        yaml: "i-vscode-icons-file-type-yaml",
        yml: "i-vscode-icons-file-type-yaml",
        toml: "i-vscode-icons-file-type-toml",
        // JS / TS ecosystem
        js: "i-vscode-icons-file-type-js",
        javascript: "i-vscode-icons-file-type-js",
        mjs: "i-vscode-icons-file-type-js",
        cjs: "i-vscode-icons-file-type-js",
        jsx: "i-vscode-icons-file-type-js",
        ts: "i-vscode-icons-file-type-typescript",
        typescript: "i-vscode-icons-file-type-typescript",
        tsx: "i-vscode-icons-file-type-typescript",
        vue: "i-vscode-icons-file-type-vue",
        coffeescript: "i-vscode-icons-file-type-coffeescript",
        // Shell
        bash: "i-vscode-icons-file-type-shell",
        sh: "i-vscode-icons-file-type-shell",
        shell: "i-vscode-icons-file-type-shell",
        zsh: "i-vscode-icons-file-type-shell",
        batch: "i-vscode-icons-file-type-shell",
        // General-purpose languages
        python: "i-vscode-icons-file-type-python",
        py: "i-vscode-icons-file-type-python",
        go: "i-vscode-icons-file-type-go",
        rust: "i-vscode-icons-file-type-rust",
        rs: "i-vscode-icons-file-type-rust",
        ruby: "i-vscode-icons-file-type-ruby",
        rb: "i-vscode-icons-file-type-ruby",
        java: "i-vscode-icons-file-type-java",
        cs: "i-vscode-icons-file-type-csharp",
        csharp: "i-vscode-icons-file-type-csharp",
        kotlin: "i-vscode-icons-file-type-kotlin",
        kt: "i-vscode-icons-file-type-kotlin",
        php: "i-vscode-icons-file-type-php",
        sql: "i-vscode-icons-file-type-sql",
        // Infra / config
        docker: "i-vscode-icons-file-type-docker",
        dockerfile: "i-vscode-icons-file-type-docker",
        // Plain text fallback
        text: "i-vscode-icons-file-type-text",
        plaintext: "i-vscode-icons-file-type-text",
      },
    },
  },
});
