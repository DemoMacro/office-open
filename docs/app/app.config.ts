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
});
