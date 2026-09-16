---
prose: true
seo:
    title: Generate Office Open XML documents with JavaScript/TypeScript
    description: Generate, parse, and patch .docx, .pptx, and .xlsx files with JSON or TypeScript — AI-native, fully typed, 100% OOXML coverage. Runs in Node.js, browsers, Deno, and Bun.
---

::page-hero{orientation="horizontal"}
#body
  :::api-example{type="docx"}

  ```json [JSON]
{
    "sections": [
        {
            "children": [
                { "paragraph": { "children": [{ "text": "Hello, World!", "bold": true }] } }
            ]
        }
    ]
}
```

  ```bash [pnpm]
pnpm add office-open
```

  ```bash [npm]
npm install office-open
```

  ```bash [yarn]
yarn add office-open
```

  ```bash [bun]
bun add office-open
```

  :::

#title
Generate Office Open XML documents.

#description
Create `.docx`, `.pptx`, and `.xlsx` files from plain JSON or fully typed TypeScript — a natural fit for AI agents and hand-written code alike.

#links
  :::button-link{to="/en/getting-started/installation"}
  Get Started <Icon name="i-lucide-arrow-right" />
  :::

  :::button-link{to="https://github.com/DemoMacro/office-open" target="_blank" variant="outline"}
  GitHub <Icon name="i-simple-icons-github" />
  :::
::

::page-section
#cards
  :::page-card
  #icon
  <Icon name="i-lucide-braces" />

  #title
  JSON & TypeScript

  #description
  Define documents as plain data — zero classes, zero boilerplate — with frozen JSON Schemas for tool-calling.
  :::

  :::page-card
  #icon
  <Icon name="i-lucide-layers" />

  #title
  Rich Content

  #description
  Paragraphs, tables, charts, images, SmartArt, math equations, headers, footers, and more.
  :::

  :::page-card
  #icon
  <Icon name="i-simple-icons-typescript" />

  #title
  Type-safe

  #description
  Comprehensive TypeScript definitions power autocomplete and catch errors as you type.
  :::

  :::page-card
  #icon
  <Icon name="i-lucide-monitor" />

  #title
  Cross-platform

  #description
  Runs in Node.js, browsers, Deno, and Bun; export to Buffer, Blob, Base64, stream, or string.
  :::

  :::page-card
  #icon
  <Icon name="i-lucide-shield-check" />

  #title
  OOXML Complete

  #description
  Every OOXML Transitional element and attribute, both generating and parsing — output opens in every major office suite.
  :::

  :::page-card
  #icon
  <Icon name="i-lucide-package" />

  #title
  Modular Packages

  #description
  Install just the format you need, or the unified package with CLI and AI SDK tools on top.
  :::
::

::page-section{orientation="horizontal"}
#body
  :::api-example

  ```json [DOCX]
{
    "sections": [
        {
            "children": [
                {
                    "table": {
                        "rows": [
                            { "cells": [{ "children": [{ "paragraph": "Name" }] }, { "children": [{ "paragraph": "Role" }] }] },
                            { "cells": [{ "children": [{ "paragraph": "Alice" }] }, { "children": [{ "paragraph": "Engineer" }] }] },
                            { "cells": [{ "children": [{ "paragraph": "Bob" }] }, { "children": [{ "paragraph": "Designer" }] }] }
                        ]
                    }
                }
            ]
        }
    ]
}
```

  ```json [PPTX]
{
    "slides": [
        {
            "children": [
                {
                    "shape": {
                        "x": 100, "y": 100, "width": 760, "height": 340,
                        "textBody": { "paragraphs": [{ "children": [{ "text": "Hello, World!", "size": 32 }] }] }
                    }
                }
            ]
        }
    ]
}
```

  ```json [XLSX]
{
    "worksheets": [
        {
            "name": "Sheet1",
            "rows": [
                { "cells": [{ "value": "Name" }, { "value": "Score" }] },
                { "cells": [{ "value": "Alice" }, { "value": 95 }] },
                { "cells": [{ "value": "Bob" }, { "value": 88 }] }
            ]
        }
    ]
}
```

  :::

#title
Build documents with JSON or TypeScript

#description
Define documents as plain JSON objects, or reach for the TypeScript API for a full IDE experience. Both produce the same valid OOXML markup.

#links
  :::button-link{to="/en/docx/quickstart" variant="outline" size="sm"}
  Explore Word docs <Icon name="i-lucide-arrow-right" />
  :::

  :::button-link{to="/en/pptx/quickstart" variant="outline" size="sm"}
  Explore PowerPoint <Icon name="i-lucide-arrow-right" />
  :::

  :::button-link{to="/en/xlsx/quickstart" variant="outline" size="sm"}
  Explore Excel <Icon name="i-lucide-arrow-right" />
  :::

#features
  :::page-feature
  #icon
  <Icon name="i-lucide-file-text" />

  #title
  Create Word documents with paragraphs, tables, images, and charts
  :::

  :::page-feature
  #icon
  <Icon name="i-lucide-presentation" />

  #title
  Create PowerPoint presentations with shapes, animations, and transitions
  :::

  :::page-feature
  #icon
  <Icon name="i-lucide-table-2" />

  #title
  Create Excel spreadsheets with styles, charts, and data validation
  :::

  :::page-feature
  #icon
  <Icon name="i-lucide-zap" />

  #title
  High Performance with native zlib compression and streaming output
  :::
::

::page-section{orientation="horizontal" reverse}
#body
  ::code-group

  ```ts [DOCX]
import { parseDocument, patchDocument } from "@office-open/docx";

// Parse existing file
const opts = await parseDocument(buffer);
// opts.sections — document sections
// opts.title, opts.creator — core properties

// Patch template placeholders
const result = await patchDocument({
  outputType: "nodebuffer",
  data: buffer,
  placeholders: {
    name: { type: "paragraph", children: [{ text: "John" }] },
  },
});
```

  ```ts [PPTX]
import { parsePresentation, patchPresentation } from "@office-open/pptx";

// Parse existing file
const opts = await parsePresentation(buffer);
// opts.slides — slide array
// opts.size, opts.title — presentation properties

// Patch template placeholders
const result = await patchPresentation({
  outputType: "nodebuffer",
  data: buffer,
  placeholders: {
    title: [{ text: "Updated", bold: true }],
  },
});
```

  ```ts [XLSX]
import { parseWorkbook, patchWorkbook } from "@office-open/xlsx";

// Parse existing file
const opts = await parseWorkbook(buffer);
// opts.worksheets — worksheet array
// opts.styles — style definitions

// Patch template placeholders
const result = await patchWorkbook({
  outputType: "nodebuffer",
  data: buffer,
  placeholders: {
    name: "John Doe",
  },
});
```

  ::

#title
Read and modify existing files

#description
Parse `.docx`, `.pptx`, and `.xlsx` files into structured objects, or patch templates by replacing `{{placeholder}}` tokens with new content.

#links
  :::button-link{to="/en/docx/parsing" variant="outline" size="sm"}
  Parse documents <Icon name="i-lucide-arrow-right" />
  :::

  :::button-link{to="/en/docx/patch" variant="outline" size="sm"}
  Patch templates <Icon name="i-lucide-arrow-right" />
  :::

#features
  :::page-feature
  #icon
  <Icon name="i-lucide-search" />

  #title
  Read document structure, styles, and content
  :::

  :::page-feature
  #icon
  <Icon name="i-lucide-wrench" />

  #title
  Patch template placeholders with new content
  :::

  :::page-feature
  #icon
  <Icon name="i-lucide-arrow-right-left" />

  #title
  Parse, modify, and re-export in a pipeline
  :::
::

::page-section
#title
Add document generation to your project.

#links
  :::button-link{to="/en/getting-started/installation"}
  Get Started <Icon name="i-lucide-arrow-right" />
  :::
::
