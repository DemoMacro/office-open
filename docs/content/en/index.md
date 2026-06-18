---
prose: true
seo:
    title: Generate Office Open XML documents with JavaScript/TypeScript
    description: Create .docx, .pptx, and .xlsx files programmatically with a declarative API. Works in Node.js and browsers.
---

::u-page-hero
---
orientation: horizontal
---

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
Create `.docx`, `.pptx`, and `.xlsx` files with JSON or TypeScript. Ideal for AI agents and traditional workflows alike.

#links
:::u-button
---
label: Get Started
size: lg
to: /en/getting-started/installation
trailingIcon: i-lucide-arrow-right
---
:::

:::u-button
---
label: GitHub
icon: i-simple-icons-github
size: lg
target: _blank
to: https://github.com/DemoMacro/office-open
variant: outline
---
:::
::

::u-container{.pb-12.xl:pb-24}
:::u-page-grid
::::u-page-feature
---
icon: i-lucide-braces
---
#title{unwrap="p"}
JSON & TypeScript

#description{unwrap="p"}
Create documents with pure JSON objects or the TypeScript functional API. JSON-first design makes it ideal for AI agents and LLM workflows.
::::

::::u-page-feature
---
icon: i-lucide-layers
---
#title{unwrap="p"}
Rich Content

#description{unwrap="p"}
Paragraphs, tables, images, charts, SmartArt, math equations, headers, footers, and more.
::::

::::u-page-feature
---
icon: i-simple-icons-typescript
---
#title{unwrap="p"}
Type-safe

#description{unwrap="p"}
Complete TypeScript definitions with full autocomplete and type safety.
::::

::::u-page-feature
---
icon: i-lucide-monitor
---
#title{unwrap="p"}
Cross-platform

#description{unwrap="p"}
Works in Node.js and browsers. Export to Buffer, Blob, Base64, stream, or string.
::::

::::u-page-feature
---
icon: i-lucide-shield-check
---
#title{unwrap="p"}
OOXML Compliant

#description{unwrap="p"}
Generates files that fully comply with the ISO/IEC 29500 Office Open XML specification.
::::

::::u-page-feature
---
icon: i-lucide-package
---
#title{unwrap="p"}
Modular Packages

#description{unwrap="p"}
Install only what you need — docx, pptx, xml, or core.
::::
:::
::

::u-page-section
---
orientation: horizontal
---

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
                        "textBody": { "children": [{ "text": "Hello, World!", "size": 32 }] }
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
Build documents with [JSON]{.text-(--ui-primary)} or [TypeScript]{.text-(--ui-primary)}

#description
Define documents as plain JSON objects — perfect for AI agents — or use the TypeScript functional API for a full IDE experience. Both produce valid OOXML markup.

#features
:::u-page-feature
---
icon: i-lucide-file-text
---
#title{unwrap="p"}
Create Word documents with paragraphs, tables, images, and charts
:::

:::u-page-feature
---
icon: i-lucide-presentation
---
#title{unwrap="p"}
Create PowerPoint presentations with shapes, animations, and transitions
:::

:::u-page-feature
---
icon: i-lucide-table-2
---
#title{unwrap="p"}
Create Excel spreadsheets with styles, charts, and data validation
:::

:::u-page-feature
---
icon: i-lucide-download
---
#title{unwrap="p"}
Export to Buffer, Blob, Base64, stream, or string
:::

#links
:::u-button
---
color: neutral
label: Explore @office-open/docx
to: /en/docx/quickstart
trailingIcon: i-lucide-arrow-right
variant: subtle
---
:::

:::u-button
---
color: neutral
label: Explore @office-open/pptx
to: /en/pptx/quickstart
trailingIcon: i-lucide-arrow-right
variant: subtle
---
:::
::

::u-page-section
---
orientation: horizontal
reverse: true
---

:::code-group

```ts [DOCX]
import { parseDocument, patchDocument, PatchType } from "@office-open/docx";

// Parse existing file
const opts = parseDocument(buffer);
// opts.sections — document sections
// opts.title, opts.creator — core properties

// Patch template placeholders
const result = await patchDocument({
  outputType: "nodebuffer",
  data: buffer,
  patches: {
    name: { type: PatchType.PARAGRAPH, children: [{ text: "John" }] },
  },
});
```

```ts [PPTX]
import { parsePresentation, patchPresentation, PatchType } from "@office-open/pptx";

// Parse existing file
const opts = parsePresentation(buffer);
// opts.slides — slide array
// opts.size, opts.title — presentation properties

// Patch template placeholders
const result = await patchPresentation({
  outputType: "nodebuffer",
  data: buffer,
  patches: {
    title: { type: PatchType.PARAGRAPH, children: [{ text: "Updated", bold: true }] },
  },
});
```

```ts [XLSX]
import { parseWorkbook, patchWorkbook } from "@office-open/xlsx";

// Parse existing file
const opts = parseWorkbook(buffer);
// opts.worksheets — worksheet array
// opts.styles — style definitions

// Patch template placeholders
const result = await patchWorkbook({
  outputType: "nodebuffer",
  data: buffer,
  patches: {
    name: { value: "John Doe" },
  },
});
```

:::

#title
Read and [modify]{.text-(--ui-primary)} existing files

#description
Parse `.docx`, `.pptx`, and `.xlsx` files into structured objects for inspection, or patch templates by replacing `{{placeholder}}` tokens with new content.

#features
:::u-page-feature
---
icon: i-lucide-search
---
#title{unwrap="p"}
Read document structure, styles, and content
:::

:::u-page-feature
---
icon: i-lucide-wrench
---
#title{unwrap="p"}
Patch template placeholders with new content
:::

:::u-page-feature
---
icon: i-lucide-arrow-right-left
---
#title{unwrap="p"}
Parse, modify, and re-export in a pipeline
:::

#links
:::u-button
---
color: neutral
label: Explore @office-open/xlsx
to: /en/xlsx/quickstart
trailingIcon: i-lucide-arrow-right
variant: subtle
---
:::
::

::u-page-section
#title
Add document generation to your project.

#links
:u-button{label="Get Started" to="/en/getting-started/installation" trailing-icon="i-lucide-arrow-right"}
::
