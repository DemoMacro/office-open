# docx-plus

![npm version](https://img.shields.io/npm/v/docx-plus)
![npm downloads](https://img.shields.io/npm/dw/docx-plus)
![npm license](https://img.shields.io/npm/l/docx-plus)

Easily generate and modify .docx files with JS/TS. Works for Node.js and on the Browser.

**docx-plus** is an enhanced fork of [docx](https://github.com/dolanmiu/docx) by Dolan Miu — a TypeScript/JavaScript library for generating and modifying Word documents (.docx) programmatically with a declarative API. We extend our sincere gratitude to the original author and all contributors for their foundational work.

This package re-exports everything from `@office-open/docx` for backward compatibility. If you are starting a new project, we recommend using `@office-open/docx` directly.

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
| OOXML compliance        | ECMA-376 (legacy)                                                          | **ISO/IEC 29500-4** (latest) — comprehensive WordprocessingML, DrawingML, and Shared Math coverage                                                |

## OOXML Coverage

**docx-plus** provides comprehensive coverage of the WordprocessingML, DrawingML, and Shared Math specifications within ISO/IEC 29500-4. This includes:

- **WordprocessingML**: paragraphs, runs, tables (full row/cell/table properties), sections, headers/footers, footnotes/endnotes, table of contents, numbering/lists, styles, bookmarks, hyperlinks, SDT content controls, form fields, track changes, comments, bibliography sources, math equations, ruby annotations, and more
- **DrawingML**: shapes, images, text bodies, colors (all 6 color types + 29 transforms), fills (solid, gradient, pattern, group), outlines, effects (28 types with recursive containers), 3D scenes, custom geometry, image adjustments, **charts** (bar, column, line, pie, area, scatter), and **SmartArt** diagrams (process, hierarchy, cycle, pyramid, list)
- **Shared Math**: all equation structures (fraction, radical, n-ary, integrals, matrices, accents, delimiters, etc.)

**Not in scope** (separate OOXML domains with their own dedicated specifications):

- **Mail Merge** — field codes are supported, but external data source connections are not
- **Cover Pages** — `docPartObj` SDT is supported, but Building Blocks part references require template files

## Installation

```bash
# Install with npm
$ npm install docx-plus

# Install with pnpm
$ pnpm add docx-plus
```

## Quick Start

```typescript
import { Document, Paragraph, TextRun, Packer } from "docx-plus";
import { writeFileSync } from "node:fs";

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
writeFileSync("My Document.docx", buffer);
```

## Examples

Check the [demo folder](../docx/demo) for 100+ working examples covering every feature.

## License

- [MIT](LICENSE) &copy; [Demo Macro](https://imst.xyz/)
