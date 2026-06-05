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
        json: "i-vscode-icons-file-type-json",
        class: "i-vscode-icons-file-type-typescript",
        docx: "i-vscode-icons-file-type-word",
        pptx: "i-vscode-icons-file-type-powerpoint",
        xlsx: "i-vscode-icons-file-type-excel",
      },
    },
  },
});
