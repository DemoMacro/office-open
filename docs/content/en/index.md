---
prose: true
seo:
    title: Generate Office Open XML documents with JavaScript/TypeScript
    description: Create .docx, .pptx, and .xlsx files programmatically with a declarative API. Works in Node.js and browsers.
---

## ::u-page-hero

## orientation: horizontal

:::code-group

```ts [example.ts]
import { Document, Packer, Paragraph, TextRun } from "@office-open/docx";

const doc = new Document({
    sections: [
        {
            children: [
                new Paragraph({
                    children: [new TextRun({ text: "Hello, World!", bold: true })],
                }),
            ],
        },
    ],
});

const buffer = await Packer.toBuffer(doc);
```

```bash [Terminal]
pnpm add @office-open/docx
```

:::

#title
Generate Office Open XML :br documents.

#description
Create `.docx`, `.pptx`, and `.xlsx` files programmatically with a declarative TypeScript API. Works in Node.js and browsers.

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
target: \_blank
to: https://github.com/DemoMacro/office-open
variant: outline

---

:::
::

::u-container{.pb-12.xl:pb-24}
:::u-page-grid
::::u-page-feature

---

## icon: i-lucide-code

#title{unwrap="p"}
Declarative API

    #description{unwrap="p"}
    Describe document structure with intuitive TypeScript classes. Each element maps to valid OOXML markup.
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
    Full type definitions and autocomplete out of the box. No additional `@types` packages needed.
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

## ::u-page-section

## orientation: horizontal

:::code-group

```ts [DOCX]
import { Document, Packer, Paragraph, TextRun } from "@office-open/docx";

const doc = new Document({
    sections: [
        {
            children: [
                new Paragraph({
                    children: [new TextRun({ text: "Hello, World!", bold: true })],
                }),
            ],
        },
    ],
});

await Packer.toBuffer(doc);
```

```ts [PPTX]
import { Presentation, Slide, Shape, Paragraph, TextRun, Packer } from "@office-open/pptx";

const pres = new Presentation({
    slides: [
        new Slide({
            children: [
                new Shape({
                    x: 1,
                    y: 1,
                    width: 8,
                    height: 4,
                    paragraphs: [
                        new Paragraph({
                            children: [new TextRun({ text: "Hello, World!", fontSize: 32 })],
                        }),
                    ],
                }),
            ],
        }),
    ],
});

await Packer.toBuffer(pres);
```

:::

#title
Build documents with [TypeScript]{.text-(--ui-primary)} classes

#description
Describe your document structure using intuitive classes. Each element produces valid OOXML markup — no XML wrestling required.

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
to: /en/docx
trailingIcon: i-lucide-arrow-right
variant: subtle

---

:::
::

## ::u-page-section

orientation: horizontal
reverse: true

---

:::code-group

```ts [DOCX]
import { parseDocx } from "@office-open/docx";

const { document, sections, paragraphs } = await parseDocx(buffer);
```

```ts [PPTX]
import { parsePptx } from "@office-open/pptx";

const { slides, shapes } = await parsePptx(buffer);
```

:::

#title
Read and [inspect]{.text-(--ui-primary)} existing files

#description
Parse existing `.docx` and `.pptx` files into structured objects. Inspect document content, extract data, or use as a foundation for modifications.

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

icon: i-lucide-file-output

---

#title{unwrap="p"}
Support for both Node.js Buffer and browser File
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
label: Explore @office-open/pptx
to: /en/pptx
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
