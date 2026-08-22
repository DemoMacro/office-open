/**
 * Styles — descriptor for xl/styles.xml (stringify delegates to
 * {@link Styles.serialize}; parse rebuilds the structured tables).
 *
 * @module
 */
import { parseOnOff } from "@office-open/core";
import type { CustomDescriptor, WriteContext } from "@office-open/core/descriptor";
import { attr, attrNum, findChild, stringifyElement } from "@office-open/xml";

import {
  parseAlignment,
  parseBorder,
  parseColorHex,
  parseFill,
  parseFont,
  parseProtection,
} from "./parse";
import type {
  CellStyleXfOptions,
  ColorsOptions,
  CustomCellStyleOptions,
  CustomTableStyleOptions,
  DxfOptions,
  CellFillOptions,
  FontOptions,
  IndexedColorOptions,
  IndexedXfEntry,
  NumFmtEntry,
  BorderSideOptions,
  StyleExtensionOptions,
  StylesDocOptions,
  StylesParseResult,
  TableStyleElementOptions,
  TableStyleElementType,
  TableStylesInfo,
} from "./types";

// ── Descriptor ──

export const stylesDesc: CustomDescriptor<StylesDocOptions, WriteContext, StylesParseResult> = {
  kind: "custom",

  stringify(opts, _ctx) {
    return opts.styles.serialize();
  },

  parse(el, _ctx) {
    const result: StylesParseResult = {};

    // numFmtById / fonts / fills / borders are hoisted so the cellStyleXfs
    // section can resolve index references into font/fill/border/numFmt defs.
    const numFmtById = new Map<number, string>();
    const fonts: FontOptions[] = [];
    const fills: CellFillOptions[] = [];
    const borders: BorderSideOptions[] = [];

    // numFmts
    const numFmtsEl = findChild(el, "numFmts");
    if (numFmtsEl) {
      const entries: NumFmtEntry[] = [];
      for (const nf of numFmtsEl.elements ?? []) {
        if (nf.name !== "numFmt") continue;
        const id = attrNum(nf, "numFmtId");
        const code = attr(nf, "formatCode");
        if (id !== undefined && code) {
          numFmtById.set(id, code);
          entries.push({ numFmtId: id, formatCode: code });
        }
      }
      result.numFmts = entries;
    }

    // fonts
    const fontsEl = findChild(el, "fonts");
    if (fontsEl) {
      for (const f of fontsEl.elements ?? []) {
        if (f.name !== "font") continue;
        fonts.push(parseFont(f));
      }
      result.fonts = fonts;
    }

    // fills
    const fillsEl = findChild(el, "fills");
    if (fillsEl) {
      for (const f of fillsEl.elements ?? []) {
        if (f.name !== "fill") continue;
        fills.push(parseFill(f));
      }
      result.fills = fills;
    }

    // borders
    const bordersEl = findChild(el, "borders");
    if (bordersEl) {
      for (const b of bordersEl.elements ?? []) {
        if (b.name !== "border") continue;
        borders.push(parseBorder(b));
      }
      result.borders = borders;
    }

    // cellStyleXfs — resolve indices into font/fill/border/numFmt definitions.
    // nativeTypeAttributes (xlsx parse path) coerces "1"/"0" to numbers, so the
    // boolean applyXxx/quotePrefix/pivotButton checks use String() coercion.
    const cellStyleXfsEl = findChild(el, "cellStyleXfs");
    if (cellStyleXfsEl) {
      const xfs: CellStyleXfOptions[] = [];
      for (const xf of cellStyleXfsEl.elements ?? []) {
        if (xf.name !== "xf") continue;
        const entry: CellStyleXfOptions = {};
        const fontId = attrNum(xf, "fontId");
        const fillId = attrNum(xf, "fillId");
        const borderId = attrNum(xf, "borderId");
        const numFmtId = attrNum(xf, "numFmtId");
        if (fontId !== undefined && fontId < fonts.length) entry.font = fonts[fontId];
        if (fillId !== undefined && fillId < fills.length) entry.fill = fills[fillId];
        if (borderId !== undefined && borderId < borders.length) entry.border = borders[borderId];
        if (numFmtId !== undefined) {
          const code = numFmtById.get(numFmtId);
          if (code !== undefined) entry.numFmt = code;
        }
        const alignmentEl = findChild(xf, "alignment");
        if (alignmentEl) entry.alignment = parseAlignment(alignmentEl);
        const protectionEl = findChild(xf, "protection");
        if (protectionEl) entry.protection = parseProtection(protectionEl);
        if (parseOnOff(attr(xf, "applyNumberFormat"))) entry.applyNumberFormat = true;
        if (parseOnOff(attr(xf, "applyFont"))) entry.applyFont = true;
        if (parseOnOff(attr(xf, "applyFill"))) entry.applyFill = true;
        if (parseOnOff(attr(xf, "applyBorder"))) entry.applyBorder = true;
        if (parseOnOff(attr(xf, "applyAlignment"))) entry.applyAlignment = true;
        if (parseOnOff(attr(xf, "applyProtection"))) entry.applyProtection = true;
        if (parseOnOff(attr(xf, "quotePrefix"))) entry.quotePrefix = true;
        if (parseOnOff(attr(xf, "pivotButton"))) entry.pivotButton = true;
        xfs.push(entry);
      }
      result.cellStyleXfs = xfs;
    }

    // cellXfs
    const cellXfsEl = findChild(el, "cellXfs");
    if (cellXfsEl) {
      const xfs: IndexedXfEntry[] = [];
      for (const xf of cellXfsEl.elements ?? []) {
        if (xf.name !== "xf") continue;
        const fontId = attrNum(xf, "fontId") ?? 0;
        const fillId = attrNum(xf, "fillId") ?? 0;
        const borderId = attrNum(xf, "borderId") ?? 0;
        const numFmtId = attrNum(xf, "numFmtId") ?? 0;

        const alignmentEl = findChild(xf, "alignment");
        const alignment = alignmentEl ? parseAlignment(alignmentEl) : undefined;

        const protectionEl = findChild(xf, "protection");
        const protection = protectionEl ? parseProtection(protectionEl) : undefined;

        const style: IndexedXfEntry = {};
        if (fontId > 0) style.fontId = fontId;
        if (fillId > 0) style.fillId = fillId;
        if (borderId > 0) style.borderId = borderId;
        if (numFmtId > 0) style.numFmtId = numFmtId;
        const xfId = attrNum(xf, "xfId");
        if (xfId !== undefined && xfId > 0) style.xfId = xfId;
        if (alignment) style.alignment = alignment;
        if (protection) style.protection = protection;
        // nativeTypeAttributes (xlsx parse path) coerces "1"/"0" to numbers
        if (parseOnOff(attr(xf, "quotePrefix"))) style.quotePrefix = true;
        if (parseOnOff(attr(xf, "pivotButton"))) style.pivotButton = true;
        // apply* flags preserved verbatim — presence distinguishes a source
        // that wrote them from one that omitted them
        if (parseOnOff(attr(xf, "applyFont"))) style.applyFont = true;
        if (parseOnOff(attr(xf, "applyFill"))) style.applyFill = true;
        if (parseOnOff(attr(xf, "applyBorder"))) style.applyBorder = true;
        if (parseOnOff(attr(xf, "applyNumberFormat"))) style.applyNumberFormat = true;
        if (parseOnOff(attr(xf, "applyAlignment"))) style.applyAlignment = true;
        if (parseOnOff(attr(xf, "applyProtection"))) style.applyProtection = true;

        xfs.push(style);
      }
      result.cellXfs = xfs;
    }

    // cellStyles
    const cellStylesEl = findChild(el, "cellStyles");
    if (cellStylesEl) {
      const styles: CustomCellStyleOptions[] = [];
      for (const cs of cellStylesEl.elements ?? []) {
        if (cs.name !== "cellStyle") continue;
        const style: Partial<CustomCellStyleOptions> = {};
        if (attr(cs, "name")) style.name = attr(cs, "name");
        const xfId = attrNum(cs, "xfId");
        if (xfId !== undefined) style.xfId = xfId;
        const builtinId = attrNum(cs, "builtinId");
        if (builtinId !== undefined) style.builtinId = builtinId;
        if (parseOnOff(attr(cs, "customBuiltin"))) style.customBuiltin = true;
        if (parseOnOff(attr(cs, "hidden"))) style.hidden = true;
        const iLevel = attrNum(cs, "iLevel");
        if (iLevel !== undefined) style.iLevel = iLevel;
        styles.push(style as CustomCellStyleOptions);
      }
      result.customCellStyles = styles;
    }

    // dxfs
    const dxfsEl = findChild(el, "dxfs");
    if (dxfsEl) {
      const dxfs: DxfOptions[] = [];
      for (const dxf of dxfsEl.elements ?? []) {
        if (dxf.name !== "dxf") continue;
        const d: DxfOptions = {};
        const fontEl = findChild(dxf, "font");
        if (fontEl) d.font = parseFont(fontEl);
        const fillEl = findChild(dxf, "fill");
        if (fillEl) d.fill = parseFill(fillEl);
        const borderEl = findChild(dxf, "border");
        if (borderEl) d.border = parseBorder(borderEl);
        const numFmtEl = findChild(dxf, "numFmt");
        if (numFmtEl && attr(numFmtEl, "formatCode")) {
          const nf: { numFmtId?: number; formatCode: string } = {
            formatCode: attr(numFmtEl, "formatCode")!,
          };
          const numFmtId = attrNum(numFmtEl, "numFmtId");
          if (numFmtId !== undefined) nf.numFmtId = numFmtId;
          d.numFmt = nf;
        }
        const alignmentEl = findChild(dxf, "alignment");
        if (alignmentEl) d.alignment = parseAlignment(alignmentEl);
        const protectionEl = findChild(dxf, "protection");
        if (protectionEl) d.protection = parseProtection(protectionEl);
        dxfs.push(d);
      }
      result.dxfs = dxfs;
    }

    // tableStyles
    const tableStylesEl = findChild(el, "tableStyles");
    if (tableStylesEl?.attributes) {
      const ts: TableStylesInfo = {};
      if (attr(tableStylesEl, "count") !== undefined)
        ts.count = attrNum(tableStylesEl, "count") ?? 0;
      if (attr(tableStylesEl, "defaultTableStyle"))
        ts.defaultTableStyle = attr(tableStylesEl, "defaultTableStyle");
      if (attr(tableStylesEl, "defaultPivotStyle"))
        ts.defaultPivotStyle = attr(tableStylesEl, "defaultPivotStyle");
      const customStyles: CustomTableStyleOptions[] = [];
      for (const tse of tableStylesEl.elements ?? []) {
        if (tse.name !== "tableStyle") continue;
        const style: Partial<CustomTableStyleOptions> = {};
        if (attr(tse, "name")) style.name = attr(tse, "name");
        if (parseOnOff(attr(tse, "pivot"))) style.pivot = true;
        const elements: TableStyleElementOptions[] = [];
        for (const tsee of tse.elements ?? []) {
          if (tsee.name !== "tableStyleElement") continue;
          const elOpts: Partial<TableStyleElementOptions> = {};
          if (attr(tsee, "type")) elOpts.type = attr(tsee, "type") as TableStyleElementType;
          const dxfId = attrNum(tsee, "dxfId");
          if (dxfId !== undefined) elOpts.dxfId = dxfId;
          elements.push(elOpts as TableStyleElementOptions);
        }
        if (elements.length > 0) style.elements = elements;
        customStyles.push(style as CustomTableStyleOptions);
      }
      if (customStyles.length > 0) ts.tableStyles = customStyles;
      result.tableStylesInfo = ts;
    }

    // colors
    const colorsEl = findChild(el, "colors");
    if (colorsEl) {
      const colors: ColorsOptions = {};
      const icEl = findChild(colorsEl, "indexedColors");
      if (icEl) {
        const indexed: IndexedColorOptions[] = [];
        for (const rgb of icEl.elements ?? []) {
          if (rgb.name === "rgbColor" && attr(rgb, "rgb")) {
            indexed.push({ rgb: attr(rgb, "rgb")! });
          }
        }
        colors.indexedColors = indexed;
      }
      const mruEl = findChild(colorsEl, "mruColors");
      if (mruEl) {
        const mru: string[] = [];
        for (const c of mruEl.elements ?? []) {
          if (c.name === "color") {
            const hex = parseColorHex(c);
            if (hex) mru.push(hex);
          }
        }
        colors.mruColors = mru;
      }
      result.colors = colors;
    }

    // styleExtensions (extLst)
    const extLstEl = findChild(el, "extLst");
    if (extLstEl) {
      const exts: StyleExtensionOptions[] = [];
      for (const ext of extLstEl.elements ?? []) {
        if (ext.name !== "ext") continue;
        const uri = attr(ext, "uri");
        if (uri) {
          // Reconstruct the inner XML of the <ext> element verbatim — each
          // child serializes itself (stringify() writes the children OF the
          // element it is given, so passing the child returns "").
          const content = (ext.elements ?? []).map((e) => stringifyElement(e)).join("");
          const namespaces: Record<string, string> = {};
          for (const [name, value] of Object.entries(ext.attributes ?? {})) {
            if (name.startsWith("xmlns:") && typeof value === "string") namespaces[name] = value;
          }
          exts.push({
            uri,
            ...(Object.keys(namespaces).length > 0 ? { namespaces } : {}),
            ...(content ? { content } : {}),
          });
        }
      }
      result.styleExtensions = exts;
    }

    return result;
  },
};
