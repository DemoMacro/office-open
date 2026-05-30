# XLSX API Reference

Complete API reference for `@office-open/xlsx`. All examples show the options JSON structure. Pass these objects to constructors (e.g. `new Workbook({ ... })`).

## Workbook Structure

```
Workbook
├── worksheets: WorksheetOptions[]
│   ├── name: string
│   ├── children: RowOptions[]
│   │   ├── cells: CellOptions[]
│   │   │   ├── value: string | number | boolean | Date | null
│   │   │   └── style: StyleOptions
│   │   ├── height: number
│   │   └── hidden: boolean
│   ├── columns: ColumnOptions[]
│   ├── mergeCells: MergeCellOptions[]
│   ├── freezePanes: FreezePaneOptions
│   ├── autoFilter: string
│   ├── images: WorksheetImageOptions[]
│   ├── charts: WorksheetChartOptions[]
│   ├── dataValidations: DataValidationOptions[]
│   └── conditionalFormats: ConditionalFormatOptions[]
├── title: string
├── creator: string
├── subject: string
└── description: string
```

## Cell Values

| `value` type | XLSX mapping            | Description                  |
| ------------ | ----------------------- | ---------------------------- |
| `string`     | Shared string reference | Deduplicated in shared table |
| `number`     | Inline numeric          | Written directly in cell     |
| `boolean`    | Boolean type (`t="b"`)  | Serialized as `0` or `1`     |
| `Date`       | Serial number           | Days since 1899-12-30 epoch  |
| `null`       | Empty cell              | Omitted from XML             |

## Cell Options

```json
{
  "value": "Hello",
  "reference": "B3",
  "styleIndex": 2,
  "style": {
    "font": { "bold": true, "italic": true, "size": 14, "color": "FF0000", "fontName": "Arial" },
    "fill": { "color": "4472C4" },
    "border": {
      "top": { "style": "thin", "color": "000000" },
      "bottom": { "style": "thin", "color": "000000" },
      "left": { "style": "thin", "color": "000000" },
      "right": { "style": "thin", "color": "000000" }
    },
    "alignment": { "horizontal": "center", "vertical": "center", "wrapText": true },
    "numFmt": "#,##0.00"
  }
}
```

| Property     | Type                                          | Description                                            |
| ------------ | --------------------------------------------- | ------------------------------------------------------ |
| `value`      | `string \| number \| boolean \| Date \| null` | Cell value                                             |
| `reference`  | `string`                                      | Cell reference (e.g. `"B3"`). Auto-assigned if omitted |
| `styleIndex` | `number`                                      | Direct style index (for pre-resolved styles)           |
| `style`      | `StyleOptions`                                | Style options (resolved to index at compile time)      |

## Style Options

### Font Options

```json
{
  "bold": true,
  "italic": true,
  "underline": true,
  "strike": true,
  "size": 14,
  "color": "FF0000",
  "fontName": "Arial"
}
```

| Property    | Type      | Description           |
| ----------- | --------- | --------------------- |
| `bold`      | `boolean` | Bold formatting       |
| `italic`    | `boolean` | Italic formatting     |
| `underline` | `boolean` | Underline formatting  |
| `strike`    | `boolean` | Strikethrough         |
| `size`      | `number`  | Font size in points   |
| `color`     | `string`  | Hex color without `#` |
| `fontName`  | `string`  | Font family name      |

### Fill Options

```json
{ "color": "4472C4" }
```

```json
{ "type": "solid", "color": "4472C4" }
```

```json
{ "patternType": "darkDown", "color": "CCCCCC" }
```

| Property      | Type     | Values                   |
| ------------- | -------- | ------------------------ |
| `type`        | `string` | `"solid"` \| `"pattern"` |
| `color`       | `string` | Hex color without `#`    |
| `patternType` | `string` | OOXML pattern type name  |

### Border Options

```json
{
  "top": { "style": "thin", "color": "000000" },
  "bottom": { "style": "thin", "color": "000000" },
  "left": { "style": "thin", "color": "000000" },
  "right": { "style": "thin", "color": "000000" },
  "diagonal": { "style": "thin", "color": "000000" }
}
```

Border side `style` values: `"thin"`, `"medium"`, `"thick"`, `"dotted"`, `"dashed"`, `"hair"`, `"none"`.

### Alignment Options

```json
{
  "horizontal": "center",
  "vertical": "center",
  "wrapText": true,
  "textRotation": 45,
  "indent": 1
}
```

| Property       | Type      | Values                                                         |
| -------------- | --------- | -------------------------------------------------------------- |
| `horizontal`   | `string`  | `"left"` \| `"center"` \| `"right"` \| `"fill"` \| `"justify"` |
| `vertical`     | `string`  | `"top"` \| `"center"` \| `"bottom"`                            |
| `wrapText`     | `boolean` | Wrap text within cell                                          |
| `textRotation` | `number`  | Rotation in degrees                                            |
| `indent`       | `number`  | Indent level                                                   |

### Number Format

```json
{ "numFmt": "#,##0.00" }
{ "numFmt": "0.00%" }
{ "numFmt": "yyyy-mm-dd" }
```

Common built-in formats: `"General"`, `"0"`, `"0.00"`, `"#,##0"`, `"#,##0.00"`, `"0%"`, `"0.00%"`, `"0.00E+00"`, `"mm-dd-yy"`, `"d-mmm-yy"`, `"h:mm"`, `"h:mm:ss"`, `"m/d/yy h:mm"`, `"@"` (text). Custom format strings start at ID 164.

## Column Options

```json
{ "min": 1, "max": 1, "width": 20, "hidden": false }
```

| Property      | Type      | Description            |
| ------------- | --------- | ---------------------- |
| `min`         | `number`  | First column (1-based) |
| `max`         | `number`  | Last column (1-based)  |
| `width`       | `number`  | Column width           |
| `hidden`      | `boolean` | Hide the column        |
| `customWidth` | `boolean` | Mark width as custom   |

## Row Options

```json
{ "cells": [{ "value": "Name" }, { "value": 30 }], "height": 30, "hidden": false }
```

| Property    | Type            | Description                |
| ----------- | --------------- | -------------------------- |
| `cells`     | `CellOptions[]` | Array of cells in this row |
| `height`    | `number`        | Row height                 |
| `hidden`    | `boolean`       | Hide the row               |
| `rowNumber` | `number`        | Row number (auto-assigned) |

## Merge Cells

```json
{ "from": { "row": 1, "col": 1 }, "to": { "row": 1, "col": 4 } }
```

| Property | Type           | Description                 |
| -------- | -------------- | --------------------------- |
| `from`   | `{ row, col }` | Top-left cell (1-based)     |
| `to`     | `{ row, col }` | Bottom-right cell (1-based) |

## Freeze Panes

```json
{ "row": 1, "col": 0 }
```

| Property | Type     | Description                                                  |
| -------- | -------- | ------------------------------------------------------------ |
| `row`    | `number` | Row split position (1-based, freezes rows above)             |
| `col`    | `number` | Column split position (1-based, freezes columns to the left) |

## Auto Filter

```json
{ "autoFilter": "A1:D10" }
```

Set directly as a string cell range on worksheet options.

## Images

```json
{ "data": "<Uint8Array>", "type": "png", "col": 3, "row": 1 }
```

| Property | Type                       | Description             |
| -------- | -------------------------- | ----------------------- |
| `data`   | `Uint8Array`               | Image binary data       |
| `type`   | `"png" \| "jpeg" \| "jpg"` | Image format            |
| `col`    | `number`                   | 1-based column position |
| `row`    | `number`                   | 1-based row position    |

## Charts

```json
{
  "type": "bar",
  "title": "Revenue",
  "categories": ["Q1", "Q2", "Q3", "Q4"],
  "series": [{ "name": "2024", "values": [100, 150, 200, 180] }],
  "col": 5,
  "row": 1
}
```

Chart types are defined by `@office-open/core` (`ChartSpaceOptions`). The `col` and `row` properties position the chart on the worksheet.

| Property     | Type       | Description                                                        |
| ------------ | ---------- | ------------------------------------------------------------------ |
| `type`       | `string`   | Chart type (e.g. `"bar"`, `"column"`, `"line"`, `"pie"`, `"area"`) |
| `title`      | `string`   | Chart title                                                        |
| `categories` | `string[]` | Category labels                                                    |
| `series`     | `array`    | `{ name, values }` series data                                     |
| `col`        | `number`   | 1-based column position                                            |
| `row`        | `number`   | 1-based row position                                               |

## Data Validation

```json
{
  "type": "list",
  "sqref": "A2:A10",
  "formula1": "\"Yes,No,Maybe\"",
  "allowBlank": true,
  "showErrorMessage": true,
  "errorTitle": "Invalid Input",
  "error": "Please select Yes, No, or Maybe"
}
```

```json
{
  "type": "whole",
  "sqref": "B2:B5",
  "operator": "between",
  "formula1": "0",
  "formula2": "100",
  "showErrorMessage": true
}
```

| Property           | Type      | Description                    |
| ------------------ | --------- | ------------------------------ |
| `sqref`            | `string`  | Cell range (e.g. `"A2:A10"`)   |
| `type`             | `string`  | See values below               |
| `operator`         | `string`  | See values below               |
| `formula1`         | `string`  | First formula / value          |
| `formula2`         | `string`  | Second formula (for `between`) |
| `allowBlank`       | `boolean` | Allow blank entries            |
| `showErrorMessage` | `boolean` | Show error alert               |
| `errorTitle`       | `string`  | Error alert title              |
| `error`            | `string`  | Error alert message            |
| `showInputMessage` | `boolean` | Show input prompt              |
| `promptTitle`      | `string`  | Input prompt title             |
| `prompt`           | `string`  | Input prompt message           |

Validation types: `"none"`, `"whole"`, `"decimal"`, `"list"`, `"date"`, `"time"`, `"textLength"`, `"custom"`.

Operators: `"between"`, `"notBetween"`, `"equal"`, `"notEqual"`, `"greaterThan"`, `"lessThan"`, `"greaterThanOrEqual"`, `"lessThanOrEqual"`.

## Conditional Formatting

```json
{
  "sqref": "B2:B5",
  "rules": [
    {
      "type": "cellIs",
      "operator": "greaterThan",
      "formulas": ["100"],
      "priority": 1,
      "dxfId": 0
    }
  ]
}
```

| Property | Type     | Description                      |
| -------- | -------- | -------------------------------- |
| `sqref`  | `string` | Cell range                       |
| `rules`  | `array`  | Array of `ConditionalFormatRule` |

### Conditional Format Rule

| Property   | Type       | Description                            |
| ---------- | ---------- | -------------------------------------- |
| `type`     | `string`   | See values below                       |
| `operator` | `string`   | See values below                       |
| `formulas` | `string[]` | Formula(s), up to 3                    |
| `priority` | `number`   | Rule priority                          |
| `dxfId`    | `number`   | Reference to differential format style |

Rule types: `"cellIs"`, `"containsText"`, `"expression"`, `"top10"`, `"aboveAverage"`.

Rule operators: `"lessThan"`, `"lessThanOrEqual"`, `"equal"`, `"notEqual"`, `"greaterThanOrEqual"`, `"greaterThan"`, `"between"`, `"notBetween"`, `"containsText"`, `"notContains"`, `"beginsWith"`, `"endsWith"`.

## Core Properties

```json
{
  "title": "My Workbook",
  "creator": "Author",
  "subject": "Subject",
  "keywords": "xlsx, spreadsheet",
  "description": "Description",
  "lastModifiedBy": "Author",
  "revision": 1
}
```

| Property         | Type     | Description     |
| ---------------- | -------- | --------------- |
| `title`          | `string` | Document title  |
| `creator`        | `string` | Author          |
| `subject`        | `string` | Subject         |
| `keywords`       | `string` | Keywords        |
| `description`    | `string` | Description     |
| `lastModifiedBy` | `string` | Last modifier   |
| `revision`       | `number` | Revision number |

## Export

```ts
import { Workbook, Packer } from "@office-open/xlsx";

const wb = new Workbook({ worksheets: [...] });

// Node.js Buffer
const buffer = await Packer.toBuffer(wb);

// Browser Blob
const blob = await Packer.toBlob(wb);

// Base64 string
const base64 = await Packer.toBase64String(wb);

// Uint8Array
const uint8 = await Packer.toUint8Array(wb);

// With formatting options
const pretty = await Packer.toBuffer(wb, { prettify: true });
```

## Parsing

### parseWorkbook (high-level)

Parse a `.xlsx` file into `WorkbookOptions`, suitable for `new Workbook(parsed)`:

```ts
import { parseWorkbook } from "@office-open/xlsx";

// Accepts Uint8Array, ArrayBuffer, Buffer, number[], base64 string
const options = parseWorkbook(readFileSync("input.xlsx"));
// options.title, options.creator, options.worksheets, etc.

const wb = new Workbook(options);
const buffer = await Packer.toBuffer(wb);
```

### parseXlsx (low-level)

Parse a `.xlsx` file into a raw `XlsxDocument` for direct XML access:

```ts
import { parseXlsx } from "@office-open/xlsx";

const doc = parseXlsx(readFileSync("input.xlsx"));
// doc.workbook — xl/workbook.xml root element
// doc.worksheets — worksheet file paths
// doc.styles — xl/styles.xml root element
// doc.sharedStrings — xl/sharedStrings.xml root element
// doc.doc — ParsedArchive (access any file via doc.doc.get(path))
// doc.partRefs — { worksheets, charts, media, drawings }
```

## Patching

### patchWorkbook

Modify an existing `.xlsx` template by replacing cell placeholder values:

```ts
import { patchWorkbook, PatchType } from "@office-open/xlsx";

const result = await patchWorkbook({
  outputType: "nodebuffer",
  data: templateBuffer,
  patches: {
    number: { value: "INV-2024-001" },
    customer: { value: "Acme Corp" },
    amount: { value: "$1,500.00" },
    date: { value: "2024-12-31" },
  },
  placeholderDelimiters: { start: "{{", end: "}}" },
});
```

Placeholders are matched in shared strings and inline strings. For string replacements, the shared string value is updated in-place.

| Option                  | Type                          | Default      | Description                   |
| ----------------------- | ----------------------------- | ------------ | ----------------------------- |
| `outputType`            | `string`                      | --           | Output format                 |
| `data`                  | `Buffer \| Uint8Array \| ...` | --           | Input .xlsx file              |
| `patches`               | `Record<string, CellPatch>`   | --           | Placeholder name -> patch map |
| `placeholderDelimiters` | `{ start, end }`              | `{ {{, }} }` | Custom delimiters             |

| PatchType        | Value    | Description      |
| ---------------- | -------- | ---------------- |
| `PatchType.CELL` | `"cell"` | Cell value patch |
