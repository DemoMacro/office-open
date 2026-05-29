---
prose: true
seo:
    title: 使用 JavaScript/TypeScript 生成 Office Open XML 文档
    description: 通过声明式 API 以编程方式创建 .docx、.pptx 和 .xlsx 文件，支持 Node.js 和浏览器。
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

```bash [Terminal]
pnpm add @office-open/docx
```

:::

#title
生成 Office Open XML :br 文档。

#description
使用 JSON 或 TypeScript 创建 `.docx`、`.pptx` 和 `.xlsx` 文件。既适合 AI 代理，也适合传统开发工作流。

#links
:::u-button
---
label: 快速开始
size: lg
to: /zh/getting-started/installation
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
JSON 与 TypeScript

#description{unwrap="p"}
使用纯 JSON 对象或 TypeScript 类创建文档。JSON 优先的设计使其成为 AI 代理和 LLM 工作流的理想选择。
::::

::::u-page-feature
---
icon: i-lucide-layers
---
#title{unwrap="p"}
丰富内容

#description{unwrap="p"}
段落、表格、图片、图表、SmartArt、数学公式、页眉、页脚等。
::::

::::u-page-feature
---
icon: i-simple-icons-typescript
---
#title{unwrap="p"}
类型安全

#description{unwrap="p"}
完整的类型定义和自动补全，开箱即用。无需额外安装 `@types` 包。
::::

::::u-page-feature
---
icon: i-lucide-monitor
---
#title{unwrap="p"}
跨平台

#description{unwrap="p"}
支持 Node.js 和浏览器。可导出为 Buffer、Blob、Base64、流或字符串。
::::

::::u-page-feature
---
icon: i-lucide-shield-check
---
#title{unwrap="p"}
符合 OOXML 规范

#description{unwrap="p"}
生成的文件完全符合 ISO/IEC 29500 Office Open XML 规范。
::::

::::u-page-feature
---
icon: i-lucide-package
---
#title{unwrap="p"}
模块化包

#description{unwrap="p"}
按需安装 — docx、pptx、xml 或 core。
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
                            { "children": [{ "children": [{ "paragraph": "Name" }] }, { "children": [{ "paragraph": "Role" }] }] },
                            { "children": [{ "children": [{ "paragraph": "Alice" }] }, { "children": [{ "paragraph": "Engineer" }] }] },
                            { "children": [{ "children": [{ "paragraph": "Bob" }] }, { "children": [{ "paragraph": "Designer" }] }] }
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
                        "textBody": { "children": [{ "text": "Hello, World!", "fontSize": 32 }] }
                    }
                }
            ]
        }
    ]
}
```

:::

#title
使用 [JSON]{.text-(--ui-primary)} 或 TypeScript 构建文档

#description
将文档定义为纯 JSON 对象——非常适合 AI 代理——或使用 TypeScript 类获得完整的 IDE 体验。两者均生成有效的 OOXML 标记。

#features
:::u-page-feature
---
icon: i-lucide-file-text
---
#title{unwrap="p"}
创建 Word 文档，支持段落、表格、图片和图表
:::

:::u-page-feature
---
icon: i-lucide-presentation
---
#title{unwrap="p"}
创建 PowerPoint 演示文稿，支持形状、动画和切换效果
:::

:::u-page-feature
---
icon: i-lucide-download
---
#title{unwrap="p"}
导出为 Buffer、Blob、Base64、流或字符串
:::

#links
:::u-button
---
color: neutral
label: 了解 @office-open/docx
to: /zh/docx/quickstart
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
import { parseDocument } from "@office-open/docx";

const opts = parseDocument(buffer);
// opts.sections — 文档节
// opts.title, opts.creator — 核心属性
```

```ts [PPTX]
import { parsePresentation } from "@office-open/pptx";

const opts = parsePresentation(buffer);
// opts.slides — 幻灯片数组
// opts.size, opts.title — 演示文稿属性
```

```ts [补丁]
import { patchDocument, PatchType } from "@office-open/docx";

const result = await patchDocument({
  outputType: "nodebuffer",
  data: buffer,
  patches: {
    name: { type: PatchType.PARAGRAPH, children: [new TextRun("John")] },
  },
});
```

:::

#title
读取和[修改]{.text-(--ui-primary)}现有文件

#description
将 `.docx` 和 `.pptx` 文件解析为结构化对象进行检查，或通过替换 `{{占位符}}` 标记来修补模板。

#features
:::u-page-feature
---
icon: i-lucide-search
---
#title{unwrap="p"}
读取文档结构、样式和内容
:::

:::u-page-feature
---
icon: i-lucide-wrench
---
#title{unwrap="p"}
替换模板占位符为新内容
:::

:::u-page-feature
---
icon: i-lucide-arrow-right-left
---
#title{unwrap="p"}
解析、修改、重新导出一站式流水线
:::

#links
:::u-button
---
color: neutral
label: 了解 @office-open/pptx
to: /zh/pptx/quickstart
trailingIcon: i-lucide-arrow-right
variant: subtle
---
:::
::

::u-page-section
#title
为你的项目添加文档生成能力。

#links
:u-button{label="快速开始" to="/zh/getting-started/installation" trailing-icon="i-lucide-arrow-right"}
::
