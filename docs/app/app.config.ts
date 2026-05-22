export default defineAppConfig({
  header: {
    title: "Office Open",
  },
  seo: {
    titleTemplate: "%s - Office Open",
    title: "Office Open",
    description:
      "Generate Office Open XML documents (.docx, .pptx, .xlsx) with JavaScript/TypeScript",
  },
  github: {
    url: "https://github.com/DemoMacro/office-open",
    branch: "main",
    rootDir: "docs",
  },
  navigation: { sub: "header" },
});
