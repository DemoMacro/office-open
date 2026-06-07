/**
 * Styles descriptor for XLSX — generates xl/styles.xml.
 *
 * The Styles class is a complex accumulator (fonts, fills, borders, cellXfs
 * are registered during compilation). The descriptor reads the accumulator's
 * state via toDescriptorOptions() and produces XML inline.
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attrs, escapeXml, findChild, attr, attrNum } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";

import type {
  Styles,
  StylesState,
  FontOptions,
  FillOptions,
  BorderOptions,
  BorderSideOptions,
  AlignmentOptions,
  StyleOptions,
  CellProtectionOptions,
} from "../../file/styles";

// ── Types ──

export interface StylesDocOptions {
  /** The Styles accumulator instance (for stringify). */
  styles: Styles;
}

// ── Descriptor ──

export const stylesDesc: CustomDescriptor<StylesDocOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    return stringifyStyles(opts.styles.toDescriptorOptions());
  },

  parse(el, _ctx) {
    // Parse styles.xml into a structured result.
    // This is the read-path counterpart — extract fonts, fills, borders, cellXfs.
    const result: Record<string, unknown> = {};

    // numFmts
    const numFmtsEl = findChild(el, "numFmts");
    if (numFmtsEl) {
      const numFmts: Record<string, number> = {};
      for (const nf of numFmtsEl.elements ?? []) {
        if (nf.name !== "numFmt") continue;
        const id = attrNum(nf, "numFmtId");
        const code = attr(nf, "formatCode");
        if (id !== undefined && code) numFmts[code] = id;
      }
      result.customNumFmts = numFmts;
    }

    // fonts
    const fontsEl = findChild(el, "fonts");
    if (fontsEl) {
      const fonts: FontOptions[] = [];
      for (const f of fontsEl.elements ?? []) {
        if (f.name !== "font") continue;
        fonts.push(parseFont(f));
      }
      result.fonts = fonts;
    }

    // fills
    const fillsEl = findChild(el, "fills");
    if (fillsEl) {
      const fills: FillOptions[] = [];
      for (const f of fillsEl.elements ?? []) {
        if (f.name !== "fill") continue;
        fills.push(parseFill(f));
      }
      result.fills = fills;
    }

    // borders
    const bordersEl = findChild(el, "borders");
    if (bordersEl) {
      const borders: BorderSideOptions[] = [];
      for (const b of bordersEl.elements ?? []) {
        if (b.name !== "border") continue;
        borders.push(parseBorder(b));
      }
      result.borders = borders;
    }

    // cellXfs
    const cellXfsEl = findChild(el, "cellXfs");
    if (cellXfsEl) {
      const xfs: StyleOptions[] = [];
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

        const style: Record<string, unknown> = {};
        if (fontId > 0) style.fontIdx = fontId;
        if (fillId > 0) style.fillIdx = fillId;
        if (borderId > 0) style.borderIdx = borderId;
        if (numFmtId > 0) style.numFmtIdx = numFmtId;
        if (alignment) style.alignment = alignment;
        if (protection) style.protection = protection;
        if (attr(xf, "quotePrefix") === "1") style.quotePrefix = true;
        if (attr(xf, "pivotButton") === "1") style.pivotButton = true;

        xfs.push(style as StyleOptions);
      }
      result.cellXfs = xfs;
    }

    return result as Record<string, unknown>;
  },
};

// ── Parse helpers ──

function parseFont(el: XmlElement): FontOptions {
  const result: Record<string, unknown> = {};
  for (const child of el.elements ?? []) {
    switch (child.name) {
      case "b":
        result.bold = true;
        break;
      case "i":
        result.italic = true;
        break;
      case "u":
        result.underline = true;
        break;
      case "strike":
        result.strike = true;
        break;
      case "outline":
        result.outline = true;
        break;
      case "shadow":
        result.shadow = true;
        break;
      case "condense":
        result.condense = true;
        break;
      case "extend":
        result.extend = true;
        break;
      case "sz":
        result.size = attrNum(child, "val");
        break;
      case "color":
        result.color = parseColorHex(child);
        break;
      case "name":
        result.fontName = attr(child, "val") ?? undefined;
        break;
      case "charset":
        result.charset = attrNum(child, "val");
        break;
      case "family":
        result.family = attrNum(child, "val");
        break;
      case "vertAlign":
        result.vertAlign = (attr(child, "val") as FontOptions["vertAlign"]) ?? undefined;
        break;
      case "scheme":
        result.scheme = (attr(child, "val") as FontOptions["scheme"]) ?? undefined;
        break;
    }
  }
  return result as FontOptions;
}

function parseFill(el: XmlElement): FillOptions {
  const patternFill = findChild(el, "patternFill");
  if (patternFill) {
    const result: Record<string, unknown> = {};
    result.patternType = attr(patternFill, "patternType") ?? undefined;
    const fg = findChild(patternFill, "fgColor");
    if (fg) result.color = parseColorHex(fg);
    const bg = findChild(patternFill, "bgColor");
    if (bg) result.bgColor = parseColorHex(bg);
    const indexed = fg ? attrNum(fg, "indexed") : undefined;
    if (indexed !== undefined) result.colorIndexed = indexed;
    return result as FillOptions;
  }

  const gradientFill = findChild(el, "gradientFill");
  if (gradientFill) {
    const result: Record<string, unknown> = { type: "gradient" };
    const gType = attr(gradientFill, "type");
    if (gType) result.gradientType = gType as FillOptions["gradientType"];
    const degree = attrNum(gradientFill, "degree");
    if (degree !== undefined) result.gradientDegree = degree;
    const stops: { position: number; color: string }[] = [];
    for (const s of gradientFill.elements ?? []) {
      if (s.name !== "stop") continue;
      const pos = attrNum(s, "position");
      const color = findChild(s, "color");
      if (pos !== undefined && color) {
        stops.push({ position: pos, color: parseColorHex(color) ?? "" });
      }
    }
    if (stops.length > 0) result.stops = stops;
    return result as FillOptions;
  }

  return {} as FillOptions;
}

function parseBorder(el: XmlElement): BorderSideOptions {
  const result: Record<string, unknown> = {};
  if (attr(el, "diagonalUp") === "1") result.diagonalUp = true;
  if (attr(el, "diagonalDown") === "1") result.diagonalDown = true;

  for (const side of ["left", "right", "top", "bottom", "diagonal"] as const) {
    const sideEl = findChild(el, side);
    if (sideEl) {
      const opts: Record<string, unknown> = {};
      const style = attr(sideEl, "style");
      if (style) opts.style = style as BorderOptions["style"];
      const color = findChild(sideEl, "color");
      if (color) opts.color = parseColorHex(color);
      if (Object.keys(opts).length > 0) result[side] = opts;
    }
  }

  return result as BorderSideOptions;
}

function parseAlignment(el: XmlElement): AlignmentOptions {
  const result: Record<string, unknown> = {};
  const h = attr(el, "horizontal");
  if (h) result.horizontal = h as AlignmentOptions["horizontal"];
  const v = attr(el, "vertical");
  if (v) result.vertical = v as AlignmentOptions["vertical"];
  if (attr(el, "wrapText") === "1") result.wrapText = true;
  const rotation = attrNum(el, "textRotation");
  if (rotation !== undefined) result.textRotation = rotation;
  const indent = attrNum(el, "indent");
  if (indent !== undefined) result.indent = indent;
  return result as AlignmentOptions;
}

function parseProtection(el: XmlElement): CellProtectionOptions {
  const result: Record<string, unknown> = {};
  const locked = attr(el, "locked");
  if (locked !== undefined) result.locked = locked !== "0";
  const hidden = attr(el, "hidden");
  if (hidden !== undefined) result.hidden = hidden !== "0";
  return result as CellProtectionOptions;
}

function parseColorHex(el: XmlElement): string | undefined {
  const rgb = attr(el, "rgb");
  if (rgb) {
    // Strip alpha prefix if present (FF000000 → 000000)
    return rgb.length === 8 ? rgb.slice(2) : rgb;
  }
  return undefined;
}

// ── Stringify (inlined from Styles.toXml) ──

function stringifyStyles(s: StylesState): string {
  const p: string[] = [
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
  ];

  // numFmts
  if (s.customNumFmts.size > 0) {
    p.push(`<numFmts count="${s.customNumFmts.size}">`);
    for (const [fmt, id] of s.customNumFmts) {
      p.push(`<numFmt numFmtId="${id}" formatCode="${escapeXml(fmt)}"/>`);
    }
    p.push("</numFmts>");
  }

  // fonts
  p.push(`<fonts count="${s.fonts.length}">`);
  for (const f of s.fonts) {
    p.push(`<font>${fontXmlStr(f)}</font>`);
  }
  p.push("</fonts>");

  // fills
  p.push(`<fills count="${s.fills.length}">`);
  for (const f of s.fills) {
    if (f.type === "gradient" && f.stops && f.stops.length > 0) {
      const gfAttrs: Record<string, string | number | boolean | undefined> = {};
      if (f.gradientType && f.gradientType !== "linear") gfAttrs.type = f.gradientType;
      if (f.gradientDegree !== undefined) gfAttrs.degree = f.gradientDegree;
      if (f.gradientLeft !== undefined) gfAttrs.left = f.gradientLeft;
      if (f.gradientRight !== undefined) gfAttrs.right = f.gradientRight;
      if (f.gradientTop !== undefined) gfAttrs.top = f.gradientTop;
      if (f.gradientBottom !== undefined) gfAttrs.bottom = f.gradientBottom;
      const stopParts = f.stops
        .map((st) => `<stop position="${st.position}"><color rgb="FF${st.color}"/></stop>`)
        .join("");
      p.push(`<fill><gradientFill${attrs(gfAttrs)}>${stopParts}</gradientFill></fill>`);
    } else {
      const patternAttrs = attrs({ patternType: f.patternType ?? "solid" });
      const fgColor = f.color
        ? `<fgColor rgb="FF${f.color}"/>`
        : f.colorIndexed !== undefined
          ? `<fgColor indexed="${f.colorIndexed}"/>`
          : "";
      const bgColor = f.bgColor ? `<bgColor rgb="FF${f.bgColor}"/>` : "";
      const colorContent = fgColor + bgColor;
      p.push(
        colorContent
          ? `<fill><patternFill${patternAttrs}>${colorContent}</patternFill></fill>`
          : `<fill><patternFill${patternAttrs}/></fill>`,
      );
    }
  }
  p.push("</fills>");

  // borders
  p.push(`<borders count="${s.borders.length}">`);
  for (const b of s.borders) {
    const bAttrs: string[] = [];
    if (b.diagonalUp) bAttrs.push('diagonalUp="1"');
    if (b.diagonalDown) bAttrs.push('diagonalDown="1"');
    const bAttr = bAttrs.length ? ` ${bAttrs.join(" ")}` : "";
    p.push(`<border${bAttr}>${borderXmlStr(b)}</border>`);
  }
  p.push("</borders>");

  // cellStyleXfs
  p.push(
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>',
  );

  // cellXfs
  p.push(`<cellXfs count="${s.cellXfs.length}">`);
  for (const xf of s.cellXfs) {
    const xAttrs: Record<string, string | number | boolean | undefined> = {
      numFmtId: xf.numFmtId,
      fontId: xf.fontId,
      fillId: xf.fillId,
      borderId: xf.borderId,
      xfId: 0,
    };
    if (xf.alignment) xAttrs.applyAlignment = 1;
    if (xf.fontId > 0) xAttrs.applyFont = 1;
    if (xf.fillId > 0) xAttrs.applyFill = 1;
    if (xf.borderId > 0) xAttrs.applyBorder = 1;
    if (xf.numFmtId > 0) xAttrs.applyNumberFormat = 1;
    if (xf.quotePrefix) xAttrs.quotePrefix = 1;
    if (xf.pivotButton) xAttrs.pivotButton = 1;
    if (xf.applyProtection) xAttrs.applyProtection = 1;
    if (xf.protection) xAttrs.applyProtection = xAttrs.applyProtection ?? 1;

    const alignStr = xf.alignment ? alignmentXmlStr(xf.alignment) : "";
    const protStr = xf.protection ? protectionXmlStr(xf.protection) : "";
    const inner = alignStr + protStr;
    p.push(inner ? `<xf${attrs(xAttrs)}>${inner}</xf>` : `<xf${attrs(xAttrs)}/>`);
  }
  p.push("</cellXfs>");

  // cellStyles
  if (s.customCellStyles && s.customCellStyles.length > 0) {
    const csAttrs: string[] = [`count="${s.customCellStyles.length + 1}"`];
    const csParts: string[] = [`<cellStyles ${csAttrs.join(" ")}>`];
    csParts.push('<cellStyle name="Normal" xfId="0" builtinId="0"/>');
    for (const cs of s.customCellStyles) {
      const csAttrParts: string[] = [`name="${escapeXml(cs.name)}"`, `xfId="${cs.xfId}"`];
      if (cs.builtinId !== undefined) csAttrParts.push(`builtinId="${cs.builtinId}"`);
      if (cs.customBuiltin) csAttrParts.push('customBuiltin="1"');
      if (cs.iLevel !== undefined) csAttrParts.push(`iLevel="${cs.iLevel}"`);
      if (cs.hidden) csAttrParts.push('hidden="1"');
      csParts.push(`<cellStyle ${csAttrParts.join(" ")}/>`);
    }
    csParts.push("</cellStyles>");
    p.push(csParts.join(""));
  } else {
    p.push('<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>');
  }

  // dxfs
  if (s.dxfs.length > 0) {
    p.push(`<dxfs count="${s.dxfs.length}">`);
    for (const dxf of s.dxfs) {
      const dParts: string[] = [];
      if (dxf.font) dParts.push(`<font>${fontXmlStr(dxf.font)}</font>`);
      if (dxf.fill) {
        const bgColor = dxf.fill.color ? `<bgColor rgb="FF${dxf.fill.color}"/>` : "";
        const patAttrs = attrs({ patternType: dxf.fill.patternType ?? "solid" });
        dParts.push(`<fill><patternFill${patAttrs}>${bgColor}</patternFill></fill>`);
      }
      if (dxf.numFmt) dParts.push(`<numFmt formatCode="${escapeXml(dxf.numFmt)}"/>`);
      if (dxf.border) dParts.push(`<border>${borderXmlStr(dxf.border)}</border>`);
      if (dParts.length > 0) {
        p.push(`<dxf>${dParts.join("")}</dxf>`);
      } else {
        p.push("<dxf/>");
      }
    }
    p.push("</dxfs>");
  } else {
    p.push('<dxfs count="0"/>');
  }

  // tableStyles (CT_TableStyles)
  if (s.tableStyles && s.tableStyles.length > 0) {
    const tsParts: string[] = [
      `<tableStyles count="${s.tableStyles.length}" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16">`,
    ];
    for (const ts of s.tableStyles) {
      const tsAttrs: string[] = [`name="${escapeXml(ts.name)}"`];
      if (ts.pivot) tsAttrs.push('pivot="1"');
      if (ts.elements && ts.elements.length > 0) {
        tsParts.push(`<tableStyle ${tsAttrs.join(" ")}>`);
        for (const el of ts.elements) {
          const elAttrs: string[] = [`type="${el.type}"`];
          if (el.dxfId !== undefined) elAttrs.push(`dxfId="${el.dxfId}"`);
          if (el.button) elAttrs.push('button="1"');
          tsParts.push(`<tableStyleElement ${elAttrs.join(" ")}/>`);
        }
        tsParts.push("</tableStyle>");
      } else {
        tsParts.push(`<tableStyle ${tsAttrs.join(" ")}/>`);
      }
    }
    tsParts.push("</tableStyles>");
    p.push(tsParts.join(""));
  } else {
    p.push(
      '<tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>',
    );
  }

  // colors (optional color palette)
  if (s.colors) {
    const c = s.colors;
    const colorParts: string[] = ["<colors>"];
    if (c.indexedColors && c.indexedColors.length > 0) {
      colorParts.push("<indexedColors>");
      for (const ic of c.indexedColors) {
        colorParts.push(`<rgbColor rgb="${ic.rgb}"/>`);
      }
      colorParts.push("</indexedColors>");
    }
    if (c.mruColors && c.mruColors.length > 0) {
      colorParts.push("<mruColors>");
      for (const mc of c.mruColors) {
        colorParts.push(`<color rgb="FF${mc}"/>`);
      }
      colorParts.push("</mruColors>");
    }
    colorParts.push("</colors>");
    p.push(colorParts.join(""));
  }

  // extLst — style sheet extensions
  if (s.styleExtensions && s.styleExtensions.length > 0) {
    const extParts: string[] = ["<extLst>"];
    for (const ext of s.styleExtensions) {
      if (ext.content) {
        extParts.push(`<ext uri="${ext.uri}">${ext.content}</ext>`);
      } else {
        extParts.push(`<ext uri="${ext.uri}"/>`);
      }
    }
    extParts.push("</extLst>");
    p.push(extParts.join(""));
  } else {
    p.push("<extLst/>");
  }

  p.push("</styleSheet>");
  return p.join("");
}

function fontXmlStr(f: FontOptions): string {
  const parts: string[] = [];
  if (f.bold) parts.push("<b/>");
  if (f.italic) parts.push("<i/>");
  if (f.underline) parts.push("<u/>");
  if (f.strike) parts.push("<strike/>");
  if (f.outline) parts.push("<outline/>");
  if (f.shadow) parts.push("<shadow/>");
  if (f.condense) parts.push("<condense/>");
  if (f.extend) parts.push("<extend/>");
  if (f.size) parts.push(`<sz val="${f.size}"/>`);
  if (f.color) parts.push(`<color rgb="FF${f.color}"/>`);
  if (f.fontName) parts.push(`<name val="${escapeXml(f.fontName)}"/>`);
  if (f.charset !== undefined) parts.push(`<charset val="${f.charset}"/>`);
  if (f.family !== undefined) parts.push(`<family val="${f.family}"/>`);
  if (f.vertAlign) parts.push(`<vertAlign val="${f.vertAlign}"/>`);
  if (f.scheme) parts.push(`<scheme val="${f.scheme}"/>`);
  return parts.join("");
}

function borderXmlStr(b: BorderSideOptions): string {
  const parts: string[] = [];
  const renderSide = (name: string, opts: BorderOptions | undefined, required = true) => {
    if (opts && opts.style && opts.style !== "none") {
      const colorStr = opts.color ? `<color rgb="FF${opts.color}"/>` : "";
      parts.push(`<${name} style="${opts.style}">${colorStr}</${name}>`);
    } else if (required) {
      parts.push(`<${name}/>`);
    }
  };
  for (const side of [
    "left",
    "right",
    "top",
    "bottom",
    "diagonal",
    "vertical",
    "horizontal",
  ] as const) {
    renderSide(side, b[side] as BorderOptions | undefined);
  }
  renderSide("start", b.start, false);
  renderSide("end", b.end, false);
  return parts.join("");
}

function alignmentXmlStr(a: AlignmentOptions): string {
  const aAttrs: Record<string, string | number | boolean | undefined> = {};
  if (a.horizontal) aAttrs.horizontal = a.horizontal;
  if (a.vertical) aAttrs.vertical = a.vertical;
  if (a.wrapText) aAttrs.wrapText = 1;
  if (a.textRotation !== undefined) aAttrs.textRotation = a.textRotation;
  if (a.indent !== undefined) aAttrs.indent = a.indent;
  if (a.relativeIndent !== undefined) aAttrs.relativeIndent = a.relativeIndent;
  if (a.justifyLastLine) aAttrs.justifyLastLine = 1;
  if (a.shrinkToFit) aAttrs.shrinkToFit = 1;
  if (a.readingOrder !== undefined) aAttrs.readingOrder = a.readingOrder;
  return `<alignment${attrs(aAttrs)}/>`;
}

function protectionXmlStr(pr: CellProtectionOptions): string {
  const prAttrs: Record<string, string | number | boolean | undefined> = {};
  if (pr.locked !== undefined) prAttrs.locked = pr.locked ? 1 : 0;
  if (pr.hidden !== undefined) prAttrs.hidden = pr.hidden ? 1 : 0;
  return `<protection${attrs(prAttrs)}/>`;
}
