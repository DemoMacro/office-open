# docx-plus

[![NPM version][npm-image]][npm-url]
[![Downloads per month][downloads-image]][downloads-url]
[![GitHub Action Workflow Status][github-actions-workflow-image]][github-actions-workflow-url]
[![PRs Welcome][pr-image]][pr-url]

Easily generate and modify .docx files with JS/TS. Works for Node and on the Browser.

**docx-plus** is an enhanced fork of [docx](https://github.com/dolanmiu/docx) — a TypeScript/JavaScript library for generating and modifying Word documents (.docx) programmatically with a declarative API.

## What's Different from docx?

|                         | docx                                                                       | docx-plus                                                                                                                                         |
| ----------------------- | -------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------- |
| ZIP handling            | jszip                                                                      | **fflate** (faster, smaller)                                                                                                                      |
| Data conversion         | Manual env detection (`Buffer.from` etc.)                                  | **undio** (universal `toUint8Array`)                                                                                                              |
| `Packer.toStream()`     | Removed (pseudo-streaming)                                                 | **Restored** with real streaming ZIP                                                                                                              |
| Theme support           | Partial (missing `ThemeColor`/`ThemeFont`)                                 | **Full** (`ThemeColor`, `ThemeFont` enums, theme attributes on `Color`, `Underline`, `Border`, `Shading`)                                         |
| CT_Ind                  | Twips-only indent                                                          | **Character-based** (`startChars`, `endChars`, `hangingChars`, `firstLineChars`)                                                                  |
| CT_Spacing              | Missing `beforeLines`/`afterLines`                                         | **Complete** (both twips & line-unit spacing)                                                                                                     |
| CT_Border               | Missing `shadow`/`frame`/theme attrs                                       | **Complete** (`shadow`, `frame`, `themeColor`, `themeTint`, `themeShade`)                                                                         |
| CT_Shd                  | Missing theme fill/color attrs                                             | **Complete** (`themeColor`, `themeFill`, `themeTint`/`themeShade`)                                                                                |
| EG_RPrBase              | Missing `outline`/`shadow`/`webHidden`/`fitText`/`cs`/`eastAsianLayout`    | **Complete** (all spec-defined elements including `eastAsianLayout`)                                                                              |
| ST_Em                   | `dot` only                                                                 | **Complete** (`none`, `comma`, `circle`, `dot`, `underDot`)                                                                                       |
| ST_TabTlc               | Missing `heavy`                                                            | **Complete** (`none`, `dot`, `hyphen`, `middleDot`, `underscore`, `heavy`)                                                                        |
| CT_PPrBase              | Missing 9 paragraph props                                                  | **Complete** (`suppressAutoHyphens`, `adjustRightInd`, `snapToGrid`, `mirrorIndents`, East Asian typography, `textAlignment`, `textboxTightWrap`) |
| CT_SectPr               | Missing `noEndnote`/`bidi`/`rtlGutter`/`paperSrc`/`footnotePr`/`endnotePr` | **Complete** (all spec-defined elements)                                                                                                          |
| CT_SdtRun / CT_SdtBlock | Not implemented                                                            | **Complete** (`StructuredDocumentTagRun`, comboBox, dropDownList, date, text, equation, data binding)                                             |
| CT_FFData               | Basic checkbox only                                                        | **Complete** (`checkBox`, `ddList`, `textInput` with all options)                                                                                 |
| CT_Ruby                 | Not implemented                                                            | **Complete** (furigana/pinyin with all alignment options)                                                                                         |
| CT_TblStylePr           | Not implemented                                                            | **Complete** (all 13 override types with paragraph/run/table/row/cell properties)                                                                 |
| CT_CustomGeometry2D     | Not implemented                                                            | **Complete** (`custGeom` with moveTo, lineTo, arcTo, quadBezTo, cubicBezTo, close)                                                                |
| CT_Scene3D              | Not implemented                                                            | **Complete** (`scene3d` with camera, lightRig, backdrop)                                                                                          |
| CT_EffectContainer      | Not implemented                                                            | **Complete** (`effectDag` with 28 effect types, recursive nesting, `sib`/`tree` container types)                                                  |
| CT_TextBodyProperties   | 5 attributes, 1 child element                                              | **Complete** (all 19 attributes + prstTxWarp, EG_TextAutofit, scene3d, EG_Text3D)                                                                 |
| DrawingML colors        | RGB + scheme only                                                          | **Complete** (`EG_ColorChoice`: RGB, scheme, HSL, system, preset + `EG_ColorTransform` with 29 transforms)                                        |
| DrawingML fills         | Solid fill only                                                            | **Complete** (`solidFill`, `noFill`, `gradFill`, `pattFill`, `grpFill`)                                                                           |
| DrawingML outline       | Basic line (width, cap, fill)                                              | **Complete** (`a:ln` with `prstDash`, `lineJoin`, compound line, pen alignment, all fill types, line end markers)                                 |
| DrawingML effects       | Not implemented                                                            | **Complete** (`effectLst` + `effectDag`: blur, fillOverlay, glow, outerShdw, innerShdw, prstShdw, reflection, softEdge + 20 additional effects)   |
| DrawingML 3D            | Not implemented                                                            | **Complete** (`sp3d` + `scene3d`: bevel, extrusion, contour, materials, camera, lightRig, rotation)                                               |
| Image cropping          | Not implemented                                                            | **Supported** (`srcRect` with configurable `l/t/r/b` percentages)                                                                                 |
| Image fill mode         | Stretch only                                                               | **Complete** (`stretch` default, `tile` with scale/flip/alignment options)                                                                        |
| Image adjustment        | Not implemented                                                            | **Supported** (grayscale, luminance, HSL, tint, duotone, biLevel, alpha effects, colorChange, colorRepl, blur)                                    |
| Image hyperlink         | Click only (via wrapping `ExternalHyperlink`)                              | **Complete** (explicit `hyperlink.click`/`hyperlink.hover` on `altText`, with relationship registration)                                          |
| Math advanced           | Basic math only                                                            | **Complete** (box, borderBox, eqArr, groupChr, matrix, phant, accent, fraction, nAry, func, delim, bar, limLow, limUpp, sPre, sSub, sSup)         |
| Group Shape             | Basic (transform only)                                                     | **Enhanced** (fill, effects, `chOff`/`chExt` child coordinates)                                                                                   |
| Test environment        | jsdom                                                                      | happy-dom                                                                                                                                         |
| OOXML compliance        | ECMA-376 (legacy)                                                          | **ISO/IEC 29500-4** (latest) — comprehensive WordprocessingML, DrawingML, and Shared Math coverage; see scope notes below                         |

## OOXML Coverage

**docx-plus** provides comprehensive coverage of the WordprocessingML, DrawingML, and Shared Math specifications within ISO/IEC 29500-4. This includes:

- **WordprocessingML**: paragraphs, runs, tables (full row/cell/table properties), sections, headers/footers, footnotes/endnotes, table of contents, numbering/lists, styles, bookmarks, hyperlinks, SDT content controls, form fields, track changes, comments, bibliography sources, math equations, ruby annotations, and more
- **DrawingML**: shapes, images, text bodies, colors (all 6 color types + 29 transforms), fills (solid, gradient, pattern, group), outlines, effects (28 types with recursive containers), 3D scenes, custom geometry, and image adjustments
- **Shared Math**: all equation structures (fraction, radical, n-ary, integrals, matrices, accents, delimiters, etc.)

**Not in scope** (separate OOXML domains with their own dedicated specifications):

- **SmartArt** (`dgm:`) — graphical diagram engine
- **Charts** (`c:`) — native OOXML charting (typically handled by charting libraries)
- **Mail Merge** — field codes are supported, but external data source connections are not
- **Cover Pages** — `docPartObj` SDT is supported, but Building Blocks part references require template files

## Installation

```terminal
npm install --save docx-plus
```

```ts
import * as docx from "docx-plus";
// or
import { Document, Packer, Paragraph, TextRun } from "docx-plus";
```

## Quick Example

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

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync("My Document.docx", buffer);
```

## Documentation

Please refer to the [documentation](https://github.com/DemoMacro/docx-plus#readme) for details on how to use this library, examples and much more!

## Examples

Check the [demo folder](https://github.com/DemoMacro/docx-plus/tree/master/demo) for 110+ working examples covering every feature.

## Contributing

Contributions are welcome! Please feel free to submit pull requests.

---

[npm-image]: https://badge.fury.io/js/docx-plus.svg
[npm-url]: https://npmjs.org/package/docx-plus
[downloads-image]: https://img.shields.io/npm/dm/docx-plus.svg
[downloads-url]: https://npmjs.org/package/docx-plus
[github-actions-workflow-image]: https://github.com/DemoMacro/docx-plus/workflows/Default/badge.svg?branch=main
[github-actions-workflow-url]: https://github.com/DemoMacro/docx-plus/actions
[pr-image]: https://img.shields.io/badge/PRs-welcome-brightgreen.svg
[pr-url]: http://makeapullrequest.com
