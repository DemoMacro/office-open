# office-open

> **⚠️ Warning:** This project is not yet stable and may undergo significant changes before reaching version 1.0.0. We strongly advise against using it in production environments.

![GitHub](https://img.shields.io/github/license/DemoMacro/office-open)
[![Contributor Covenant](https://img.shields.io/badge/Contributor%20Covenant-2.1-4baaaa.svg)](https://www.contributor-covenant.org/version/2/1/code_of_conduct/)

Office Open XML monorepo - generate .docx, .pptx, .xlsx files with JS/TS.

## Packages

- **[@office-open/core](./packages/core/README.md)** - Shared OOXML infrastructure (WIP)
- **[@office-open/xml](./packages/xml/README.md)** - XML parsing/serialization (WIP)
- **[@office-open/docx](./packages/docx/README.md)** - Generate .docx files with a declarative API
- **[@office-open/xlsx](./packages/xlsx/README.md)** - Generate .xlsx files (WIP)
- **[@office-open/pptx](./packages/pptx/README.md)** - Generate .pptx files (WIP)

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

## Project Philosophy

This project follows core principles:

1. **OOXML Compliance**: Strict adherence to the ISO-IEC29500 OOXML specification
2. **Type Safety**: Full TypeScript support with comprehensive types
3. **Declarative API**: Simple, intuitive API for document generation
4. **Modular Design**: Shared core infrastructure across DOCX, PPTX, XLSX
5. **Performance**: Optimized for large documents and batch processing
6. **User Experience**: Simple API with powerful configuration options

## Support & Community

- [Report Issues](https://github.com/DemoMacro/office-open/issues)
- [Export Documentation](./packages/docx/README.md)

## License

This project is licensed under the MIT License - see the [LICENSE](./LICENSE) file for details.

---

Built with ❤️ by [Demo Macro](https://imst.xyz/)
