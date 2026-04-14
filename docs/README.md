# Welcome to docx-plus

## What is docx-plus?

**docx-plus** is an enhanced fork of [docx](https://github.com/dolanmiu/docx) — a TypeScript/JavaScript library for generating and modifying Word documents (.docx files) programmatically. It provides a declarative API that works seamlessly in both Node.js and browser environments.

## Why docx-plus?

- **Declarative API** - Define what you want (objects/config), not how to build it step-by-step
- **TypeScript First** - Full type definitions with IntelliSense support
- **Universal** - Works in Node.js, browsers, and serverless environments
- **Full Featured** - Tables, images, headers, footers, styles, and more

## Quick Navigation

| I want to...              | Go to...                                                              |
| ------------------------- | --------------------------------------------------------------------- |
| Get started quickly       | [Quickstart Guide](quickstart.md)                                     |
| Learn a specific feature  | [Usage Guides](#usage-guides)                                         |
| See working examples      | [Demo Files](https://github.com/DemoMacro/docx-plus/tree/master/demo) |
| Browse the API            | [API Documentation](https://docx-plus.js.org/api/)                    |
| Modify existing documents | [Patcher](usage/patcher.md)                                           |

## Installation

```terminal
npm install --save docx-plus
```

Then you can `require` or `import` as usual:

```ts
const docx = require("docx-plus");
```

```ts
import * as docx from "docx-plus";
// or
import { Document, Packer, Paragraph, TextRun } from "docx-plus";
```

## Basic Usage

```ts
import * as fs from "fs";
import { Document, Packer, Paragraph, TextRun } from "docx-plus";

const doc = new Document({
    sections: [
        {
            children: [
                new Paragraph({
                    children: [
                        new TextRun("Hello World"),
                        new TextRun({
                            text: " - Bold text",
                            bold: true,
                        }),
                    ],
                }),
            ],
        },
    ],
});

Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync("My Document.docx", buffer);
});
```

<p align="center">
<!-- cspell:disable-next-line -->
    <img alt="clippy the assistant" src="./clippy.png">
</p>

---

Made with 💖 by [DemoMacro](https://github.com/DemoMacro), based on [docx](https://github.com/dolanmiu/docx)
