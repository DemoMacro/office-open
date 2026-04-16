# Text Runs

!> TextRuns requires an understanding of [Paragraphs](usage/paragraph.md).

You can add multiple `text runs` in `Paragraphs`. This is the most verbose way of writing a `Paragraph` but it is also the most flexible:

```ts
import { Paragraph, TextRun } from "docx-plus";

const paragraph = new Paragraph({
    children: [
        new TextRun("My awesome text here for my university dissertation"),
        new TextRun("Foo Bar"),
    ],
});
```

Text objects have methods inside which changes the way the text is displayed.

## Typographical Emphasis

More info [here](https://english.stackexchange.com/questions/97081/what-is-the-typography-term-which-refers-to-the-usage-of-bold-italics-and-unde)

### Bold

```ts
const text = new TextRun({
    text: "Foo Bar",
    bold: true,
});
```

### Italics

```ts
const text = new TextRun({
    text: "Foo Bar",
    italics: true,
});
```

### Underline

Underline has a few options

#### Options

| Property | Type            | Notes    | Possible Values                                                                                                                                                           |
| -------- | --------------- | -------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| type     | `UnderlineType` | Optional | SINGLE, WORD, DOUBLE, THICK, DOTTED, DOTTEDHEAV, DASH, DASHEDHEAV, DASHLONG, DASHLONGHEAV, DOTDASH, DASHDOTHEAVY, DOTDOTDAS, DASHDOTDOTHEAVY, WAVE, WAVYHEAVY, WAVYDOUBLE |
| color    | `string`        | Optional | Color Hex values                                                                                                                                                          |

**Example:**

```ts
const text = new TextRun({
    text: "and then underlined ",
    underline: {
        type: UnderlineType.DOUBLE,
        color: "990011",
    },
});
```

To do a simple vanilla underline:

```ts
const text = new TextRun({
    text: "and then underlined ",
    underline: {},
});
```

### Emphasis Mark

```ts
const text = new TextRun({
    text: "and then emphasis mark",
    emphasisMark: {},
});
```

### Shading and Highlighting

```ts
const text = new TextRun({
    text: "shading",
    shading: {
        type: ShadingType.REVERSE_DIAGONAL_STRIPE,
        color: "00FFFF",
        fill: "FF0000",
    },
});
```

```ts
const text = new TextRun({
    text: "highlighting",
    highlight: "yellow",
});
```

See [demo/45-highlighting-text.ts](https://github.com/DemoMacro/docx-plus/blob/master/demo/45-highlighting-text.ts) for highlighting examples, and [demo/46-shading-text.ts](https://github.com/DemoMacro/docx-plus/blob/master/demo/46-shading-text.ts) for shading examples.

### Strike through

```ts
const text = new TextRun({
    text: "strike",
    strike: true,
});
```

### Double strike through

```ts
const text = new TextRun({
    text: "doubleStrike",
    doubleStrike: true,
});
```

### Superscript

```ts
const text = new TextRun({
    text: "superScript",
    superScript: true,
});
```

### Subscript

```ts
const text = new TextRun({
    text: "subScript",
    subScript: true,
});
```

### All Capitals

```ts
const text = new TextRun({
    text: "allCaps",
    allCaps: true,
});
```

### Small Capitals

```ts
const text = new TextRun({
    text: "smallCaps",
    smallCaps: true,
});
```

### Vanish and SpecVanish

You may want to hide your text in your document.

`Vanish` should affect the normal display of text, but an application may have settings to force hidden text to be displayed.

```ts
const text = new TextRun({
    text: "This text will be hidden",
    vanish: true,
});
```

`SpecVanish` was typically used to ensure that a paragraph style can be applied to a part of a paragraph, and still appear as in the Table of Contents (which in previous word processors would ignore the use of the style if it were being used as a character style).

```ts
const text = new TextRun({
    text: "This text will be hidden forever.",
    specVanish: true,
});
```

## Break

Sometimes you would want to put text underneath another line of text but inside the same paragraph.

```ts
const text = new TextRun({
    text: "break",
    break: 1,
});
```

Adding two breaks:

```ts
const text = new TextRun({
    text: "break",
    break: 2,
});
```

## East Asian Layout

East Asian typography settings for runs, including character combination, vertical text, and compression.

### Combined Characters

Display multiple characters inside a single character space with brackets:

```ts
// Combined characters with round brackets
const text = new TextRun({
    eastAsianLayout: {
        combine: true,
        combineBrackets: "round",
        id: 1,
    },
    text: "国民",
});

// Combined characters with square brackets
const text = new TextRun({
    eastAsianLayout: {
        combine: true,
        combineBrackets: "square",
        id: 2,
    },
    text: "日本語",
});
```

### Vertical Text

Render text vertically (for East Asian vertical writing modes):

```ts
const text = new TextRun({
    eastAsianLayout: {
        vert: true,
    },
    text: "縦書き",
});
```

`eastAsianLayout` supports the following properties: `id`, `combine`, `combineBrackets` (`"none"`, `"round"`, `"square"`, `"angle"`, `"curly"`), `vert`, and `vertCompress`.

#### Options

| Property        | Type      | Notes    | Description                                           |
| --------------- | --------- | -------- | ----------------------------------------------------- |
| id              | `number`  | Optional | Combined character ID                                 |
| combine         | `boolean` | Optional | Enable character combination                          |
| combineBrackets | `string`  | Optional | `"none"`, `"round"`, `"square"`, `"angle"`, `"curly"` |
| vert            | `boolean` | Optional | Render text vertically                                |
| vertCompress    | `boolean` | Optional | Compress characters in vertical text                  |
