# office-open

![GitHub](https://img.shields.io/github/license/DemoMacro/office-open)
[![Contributor Covenant](https://img.shields.io/badge/Contributor%20Covenant-2.1-4baaaa.svg)](https://www.contributor-covenant.org/version/2/1/code_of_conduct/)

> Generate, parse, and patch Office Open XML documents (.docx, .pptx, .xlsx) with a declarative TypeScript API. Works in Node.js and browsers.

## Features

- 📝 **Document Generation** — Create .docx, .pptx, and .xlsx files with a declarative JSON or TypeScript API
- 📖 **Parsing** — Read existing files into structured objects for inspection and round-trip workflows
- 🔧 **Template Patching** — Replace `{{placeholder}}` tokens in existing templates with new content
- 🎨 **Rich Content** — Paragraphs, tables, images, charts, SmartArt, math equations, and more
- ✨ **Effects & Animations** — Shadows, glow, 3D, transitions, entrance/exit animations (PPTX)
- 📐 **OOXML Compliant** — Output validates against ISO/IEC 29500 XSD schemas
- 🔒 **Type-safe** — Full TypeScript definitions and autocomplete, no `@types` needed
- 🌐 **Cross-platform** — Node.js and browsers. Export to Buffer, Blob, Base64, stream, or string
- ⚡ **Fast** — 1.5–9x faster than alternatives (see benchmarks in package READMEs)
- 🤖 **AI-friendly** — JSON-first API ideal for LLM workflows and AI agents

## Packages

| Package                                        | Version                                                | Description                                          |
| ---------------------------------------------- | ------------------------------------------------------ | ---------------------------------------------------- |
| [@office-open/docx](./packages/docx/README.md) | ![npm](https://img.shields.io/npm/v/@office-open/docx) | Word document generation, parsing, and patching      |
| [@office-open/pptx](./packages/pptx/README.md) | ![npm](https://img.shields.io/npm/v/@office-open/pptx) | PowerPoint generation, parsing, and patching         |
| [@office-open/xlsx](./packages/xlsx/README.md) | ![npm](https://img.shields.io/npm/v/@office-open/xlsx) | Spreadsheet generation, parsing, and patching        |
| [@office-open/core](./packages/core/README.md) | ![npm](https://img.shields.io/npm/v/@office-open/core) | Shared OOXML infrastructure, charts, unit converters |
| [@office-open/xml](./packages/xml/README.md)   | ![npm](https://img.shields.io/npm/v/@office-open/xml)  | Low-level XML parsing and serialization              |

## Quick Start

### DOCX Generation

```bash
# Install with npm
$ npm install @office-open/docx

# Install with pnpm
$ pnpm add @office-open/docx
```

```typescript
import { Document, Paragraph, TextRun, Packer } from "@office-open/docx";
import { writeFileSync } from "node:fs";

const doc = new Document({
  sections: [
    {
      children: [
        new Paragraph({
          children: [new TextRun({ text: "Hello World", bold: true })],
        }),
      ],
    },
  ],
});

const buffer = await Packer.toBuffer(doc);
writeFileSync("document.docx", buffer);
```

### PPTX Generation

```bash
# Install with npm
$ npm install @office-open/pptx

# Install with pnpm
$ pnpm add @office-open/pptx
```

```typescript
import { Presentation, Shape, Packer } from "@office-open/pptx";
import { writeFileSync } from "node:fs";

const pres = new Presentation({
  slides: [
    {
      children: [
        new Shape({
          textBody: { text: "Hello World" },
          fill: "4472C4",
          x: 100,
          y: 100,
          width: 600,
          height: 400,
        }),
      ],
    },
  ],
});

const buffer = await Packer.toBuffer(pres);
writeFileSync("presentation.pptx", buffer);
```

### XLSX Generation

```bash
# Install with npm
$ npm install @office-open/xlsx

# Install with pnpm
$ pnpm add @office-open/xlsx
```

```typescript
import { Workbook, Packer } from "@office-open/xlsx";
import { writeFileSync } from "node:fs";

const wb = new Workbook({
  worksheets: [
    {
      name: "Sheet1",
      children: [
        { cells: [{ value: "Name" }, { value: "Score" }] },
        { cells: [{ value: "Alice" }, { value: 95 }] },
        { cells: [{ value: "Bob" }, { value: 87 }] },
      ],
    },
  ],
});

const buffer = await Packer.toBuffer(wb);
writeFileSync("workbook.xlsx", buffer);
```

## Parse Existing Files

Read and inspect existing `.docx`, `.pptx`, and `.xlsx` files into structured objects:

```typescript
// DOCX
import { parseDocument } from "@office-open/docx";
const opts = parseDocument(buffer);
// opts.sections — document sections
// opts.title, opts.creator — core properties

// PPTX
import { parsePresentation } from "@office-open/pptx";
const opts = parsePresentation(buffer);
// opts.slides — slide array
// opts.size, opts.title — presentation properties

// XLSX
import { parseWorkbook } from "@office-open/xlsx";
const opts = parseWorkbook(buffer);
// opts.worksheets — worksheet array
// opts.worksheets[0].children — row/cell data
```

## JSON API

Define documents as plain JSON objects — perfect for AI agents:

```typescript
// PPTX via JSON
const pres = new Presentation({
  slides: [
    {
      children: [
        {
          shape: {
            x: 100,
            y: 100,
            width: 760,
            height: 340,
            textBody: { children: [{ text: "Hello, World!", fontSize: 32 }] },
          },
        },
      ],
    },
  ],
});
```

## Project Philosophy

This project follows core principles:

1. **OOXML Compliance**: Strict adherence to the ISO/IEC 29500 OOXML specification
2. **Type Safety**: Full TypeScript support with comprehensive types and autocomplete
3. **Declarative API**: Simple, intuitive API for document generation — JSON or TypeScript
4. **Modular Design**: Shared core infrastructure across DOCX, PPTX, and XLSX
5. **Performance**: Optimized for large documents and batch processing
6. **Cross-platform**: Works in Node.js and browsers. Export to Buffer, Blob, Base64, stream, or string

## Development

### Prerequisites

- **Node.js** 18.x or higher
- **pnpm** 9.x or higher (recommended package manager)
- **Git** for version control

### Getting Started

1. **Clone the repository**:

   ```bash
   git clone https://github.com/DemoMacro/office-open.git
   cd office-open
   ```

2. **Install dependencies**:

   ```bash
   pnpm install
   ```

3. **Development mode**:

   ```bash
   pnpm dev
   ```

4. **Build all packages**:

   ```bash
   pnpm build
   ```

5. **Test locally**:

   ```bash
   # Run tests
   pnpm test
   ```

### Development Commands

```bash
pnpm dev            # Development mode with watch
pnpm build          # Build all packages
pnpm test           # Run tests
pnpm check          # Lint & format
```

## Contributing

We welcome contributions! Here's how to get started:

### Quick Setup

1. **Fork the repository** on GitHub
2. **Clone your fork**:

   ```bash
   git clone https://github.com/YOUR_USERNAME/office-open.git
   cd office-open
   ```

3. **Add upstream remote**:

   ```bash
   git remote add upstream https://github.com/DemoMacro/office-open.git
   ```

4. **Install dependencies**:

   ```bash
   pnpm install
   ```

5. **Development mode**:

   ```bash
   pnpm dev
   ```

### Development Workflow

1. **Code**: Follow our project standards
2. **Test**: `pnpm build && pnpm test`
3. **Commit**: Use conventional commits (`feat:`, `fix:`, etc.)
4. **Push**: Push to your fork
5. **Submit**: Create a Pull Request to upstream repository

## Support & Community

- [Report Issues](https://github.com/DemoMacro/office-open/issues)
- [Export Documentation](./packages/docx/README.md)

## License

This project is licensed under the MIT License - see the [LICENSE](./LICENSE) file for details.

---

Built with ❤️ by [Demo Macro](https://imst.xyz/)
