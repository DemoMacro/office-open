export default defineAppConfig({
  docs: {
    name: "Office Open",
    description:
      "TypeScript toolkit for Office documents — generate, parse, and patch .docx, .pptx, .xlsx with spec-compliant OOXML output",
    site: "https://www.office-open.com",
    copyright: { label: "Demo Macro", url: "https://www.demomacro.com/" },
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
