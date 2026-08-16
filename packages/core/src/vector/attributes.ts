/**
 * VML attribute-group serialization — AG_AllCoreAttributes /
 * AG_AllShapeAttributes and their v: subgroups.
 *
 * Every shape-ish element (shape, shapetype, group, and the eight basic
 * shapes) carries the same attribute vocabulary, so the field↔attribute
 * mapping lives here once and the element serializers consume it. The office
 * attribute subgroups (o: prefix, injected by vml-officeDrawing.xsd) join the
 * same table in `office-attributes.ts`; a merged spec list is all a caller
 * passes to {@link stringifyVmlAttributes} / {@link parseVmlAttributes}.
 *
 * Reference: ISO/IEC 29500-4, vml-main.xsd, AG_CoreAttributes,
 * AG_ShapeAttributes and their compositions.
 *
 * @module
 */
import type { Element as XmlElement } from "@office-open/xml";
import { escapeXml } from "@office-open/xml";

import type { VmlShapeStyle } from "./style";

// ── Shared types ──

/** s:ST_TrueFalse — VML boolean attributes serialize as "t"/"f". */
export type VmlTrueFalse = boolean;

/** s:ST_ColorType kept verbatim — VML colors carry forms like "red" or "infoBackground [80]". */
export type VmlColor = string;

/** s:ST_TrueFalseBlank — a boolean that also admits the empty-string form. */
export type VmlTrueFalseBlank = boolean | "";

/** ST_BWMode (o:) — black-and-white rendering fallback. */
export type VmlBlackWhiteMode =
  | "color"
  | "auto"
  | "grayScale"
  | "lightGrayscale"
  | "inverseGray"
  | "grayOutline"
  | "highContrast"
  | "black"
  | "white"
  | "hide"
  | "undrawn"
  | "blackTextAndLines";

/** ST_ConnectorType (o:). */
export type VmlConnectorType = "none" | "straight" | "elbow" | "curved";

/** ST_HrAlign (o:) — horizontal-rule alignment. */
export type VmlHorizontalRuleAlign = "left" | "right" | "center";

/** ST_InsetMode (o:). */
export type VmlInsetMode = "auto" | "custom";

/** ST_ConnectType (o:). */
export type VmlConnectType = "none" | "rect" | "segments" | "custom";

/** ST_ScreenSize (o:) — the resolution presets Office writes verbatim. */
export type VmlScreenSize = "544,376" | "640,480" | "720,512" | "800,600" | "1024,768" | "1152,862";

/** ST_DiagramLayout (o:) — layout code kept in its XML string spelling. */
export type VmlDiagramLayout = "0" | "1" | "2" | "3";

/** AG_CoreAttributes (v: namespace). `style` expands via the style mini-language. */
export interface VmlCoreAttributes {
  id?: string;
  style?: VmlShapeStyle;
  href?: string;
  target?: string;
  class?: string;
  title?: string;
  alt?: string;
  /** "21600,21600" — coordinate space of the shape. */
  coordsize?: string;
  /** "0,0" — origin of the coordinate space. */
  coordorigin?: string;
  wrapcoords?: string;
  print?: VmlTrueFalse;
}

/** AG_ShapeAttributes (v: namespace). */
export interface VmlShapeAttributes {
  chromakey?: VmlColor;
  filled?: VmlTrueFalse;
  fillcolor?: VmlColor;
  opacity?: string;
  stroked?: VmlTrueFalse;
  strokecolor?: VmlColor;
  strokeweight?: string;
  insetpen?: VmlTrueFalse;
}

// ── Attribute table ──

/** How a field converts to/from its XML attribute form. */
export type VmlAttrKind = "string" | "trueFalse" | "trueFalseBlank" | "number";

export interface VmlAttrSpec {
  /** Options field name. */
  field: string;
  /** XML attribute name including its prefix ("o:title", "r:id"). */
  attr: string;
  kind: VmlAttrKind;
}

/** AG_CoreAttributes — `style` is special-cased (mini-language), not tabulated. */
export const VML_CORE_ATTRS: readonly VmlAttrSpec[] = [
  { field: "id", attr: "id", kind: "string" },
  { field: "href", attr: "href", kind: "string" },
  { field: "target", attr: "target", kind: "string" },
  { field: "class", attr: "class", kind: "string" },
  { field: "title", attr: "title", kind: "string" },
  { field: "alt", attr: "alt", kind: "string" },
  { field: "coordsize", attr: "coordsize", kind: "string" },
  { field: "coordorigin", attr: "coordorigin", kind: "string" },
  { field: "wrapcoords", attr: "wrapcoords", kind: "string" },
  { field: "print", attr: "print", kind: "trueFalse" },
];

/** AG_ShapeAttributes. */
export const VML_SHAPE_ATTRS: readonly VmlAttrSpec[] = [
  { field: "chromakey", attr: "chromakey", kind: "string" },
  { field: "filled", attr: "filled", kind: "trueFalse" },
  { field: "fillcolor", attr: "fillcolor", kind: "string" },
  { field: "opacity", attr: "opacity", kind: "string" },
  { field: "stroked", attr: "stroked", kind: "trueFalse" },
  { field: "strokecolor", attr: "strokecolor", kind: "string" },
  { field: "strokeweight", attr: "strokeweight", kind: "string" },
  { field: "insetpen", attr: "insetpen", kind: "trueFalse" },
];

// ── Office attribute groups (o:, vml-officeDrawing.xsd) ──

/**
 * AG_OfficeCoreAttributes — the o: members folded into AG_AllCoreAttributes,
 * carried by every shape-ish element alongside {@link VmlCoreAttributes}.
 */
export interface VmlOfficeCoreAttributes {
  /** o:spid — shape id from the drawing part's id map. */
  spid?: string;
  /** o:oned — the shape is a one-dimensional connector. */
  oned?: VmlTrueFalse;
  regroupid?: number;
  doubleclicknotify?: VmlTrueFalse;
  /** o:button — the shape behaves as a button. */
  button?: VmlTrueFalse;
  userhidden?: VmlTrueFalse;
  /** o:bullet — paragraph mark rendered as a list bullet. */
  bullet?: VmlTrueFalse;
  /** o:hr — horizontal rule. */
  hr?: VmlTrueFalse;
  hrstd?: VmlTrueFalse;
  hrnoshade?: VmlTrueFalse;
  /** Horizontal-rule width in percent. */
  hrpct?: number;
  hralign?: VmlHorizontalRuleAlign;
  allowincell?: VmlTrueFalse;
  allowoverlap?: VmlTrueFalse;
  userdrawn?: VmlTrueFalse;
  bordertopcolor?: VmlColor;
  borderleftcolor?: VmlColor;
  borderbottomcolor?: VmlColor;
  borderrightcolor?: VmlColor;
  dgmlayout?: VmlDiagramLayout;
  dgmnodekind?: number;
  dgmlayoutmru?: VmlDiagramLayout;
  insetmode?: VmlInsetMode;
}

/**
 * AG_OfficeShapeAttributes — the o: members folded into AG_AllShapeAttributes.
 */
export interface VmlOfficeShapeAttributes {
  /** o:spt — shape-type number (202 = textbox, 75 = picture frame, …). */
  spt?: number;
  connectortype?: VmlConnectorType;
  bwmode?: VmlBlackWhiteMode;
  bwpure?: VmlBlackWhiteMode;
  bwnormal?: VmlBlackWhiteMode;
  /** o:forcedash — always render dashed (accessibility). */
  forcedash?: VmlTrueFalse;
  oleicon?: VmlTrueFalse;
  /** o:ole — OLE behavior flag; ST_TrueFalseBlank ("" | boolean). */
  ole?: VmlTrueFalseBlank;
  preferrelative?: VmlTrueFalse;
  cliptowrap?: VmlTrueFalse;
  clip?: VmlTrueFalse;
}

/** AG_OfficeCoreAttributes — field names mirror the o: attribute names. */
export const VML_OFFICE_CORE_ATTRS: readonly VmlAttrSpec[] = [
  { field: "spid", attr: "o:spid", kind: "string" },
  { field: "oned", attr: "o:oned", kind: "trueFalse" },
  { field: "regroupid", attr: "o:regroupid", kind: "number" },
  { field: "doubleclicknotify", attr: "o:doubleclicknotify", kind: "trueFalse" },
  { field: "button", attr: "o:button", kind: "trueFalse" },
  { field: "userhidden", attr: "o:userhidden", kind: "trueFalse" },
  { field: "bullet", attr: "o:bullet", kind: "trueFalse" },
  { field: "hr", attr: "o:hr", kind: "trueFalse" },
  { field: "hrstd", attr: "o:hrstd", kind: "trueFalse" },
  { field: "hrnoshade", attr: "o:hrnoshade", kind: "trueFalse" },
  { field: "hrpct", attr: "o:hrpct", kind: "number" },
  { field: "hralign", attr: "o:hralign", kind: "string" },
  { field: "allowincell", attr: "o:allowincell", kind: "trueFalse" },
  { field: "allowoverlap", attr: "o:allowoverlap", kind: "trueFalse" },
  { field: "userdrawn", attr: "o:userdrawn", kind: "trueFalse" },
  { field: "bordertopcolor", attr: "o:bordertopcolor", kind: "string" },
  { field: "borderleftcolor", attr: "o:borderleftcolor", kind: "string" },
  { field: "borderbottomcolor", attr: "o:borderbottomcolor", kind: "string" },
  { field: "borderrightcolor", attr: "o:borderrightcolor", kind: "string" },
  { field: "dgmlayout", attr: "o:dgmlayout", kind: "string" },
  { field: "dgmnodekind", attr: "o:dgmnodekind", kind: "number" },
  { field: "dgmlayoutmru", attr: "o:dgmlayoutmru", kind: "string" },
  { field: "insetmode", attr: "o:insetmode", kind: "string" },
];

/** AG_OfficeShapeAttributes. */
export const VML_OFFICE_SHAPE_ATTRS: readonly VmlAttrSpec[] = [
  { field: "spt", attr: "o:spt", kind: "number" },
  { field: "connectortype", attr: "o:connectortype", kind: "string" },
  { field: "bwmode", attr: "o:bwmode", kind: "string" },
  { field: "bwpure", attr: "o:bwpure", kind: "string" },
  { field: "bwnormal", attr: "o:bwnormal", kind: "string" },
  { field: "forcedash", attr: "o:forcedash", kind: "trueFalse" },
  { field: "oleicon", attr: "o:oleicon", kind: "trueFalse" },
  { field: "ole", attr: "o:ole", kind: "string" },
  { field: "preferrelative", attr: "o:preferrelative", kind: "trueFalse" },
  { field: "cliptowrap", attr: "o:cliptowrap", kind: "trueFalse" },
  { field: "clip", attr: "o:clip", kind: "trueFalse" },
];

// ── Value coercion ──

/** s:ST_TrueFalse output form — "t"/"f", the spelling Office itself writes. */
export function stringifyVmlTrueFalse(value: boolean): "t" | "f" {
  return value ? "t" : "f";
}

/** Accepts every ST_TrueFalse input spelling. */
export function parseVmlTrueFalse(value: string): boolean {
  return value === "t" || value === "true" || value === "1";
}

// ── Table-driven attribute (de)serialization ──

/**
 * Serialize the tabulated fields of `opts` into one attribute string (leading
 * space included when non-empty). The `style` field of
 * {@link VmlCoreAttributes} joins here so callers need a single call.
 */
export function stringifyVmlAttributes(
  opts: Record<string, unknown>,
  specs: readonly VmlAttrSpec[],
): string {
  const parts: string[] = [];
  for (const spec of specs) {
    const value = opts[spec.field];
    if (value === undefined) continue;
    const text =
      spec.kind === "trueFalse"
        ? stringifyVmlTrueFalse(value as boolean)
        : spec.kind === "trueFalseBlank"
          ? stringifyVmlTrueFalseBlank(value as VmlTrueFalseBlank)
          : spec.kind === "number"
            ? String(value as number)
            : escapeXml(String(value as string));
    parts.push(`${spec.attr}="${text}"`);
  }
  return parts.length > 0 ? ` ${parts.join(" ")}` : "";
}

/** Serialize a boolean-or-empty ST_TrueFalseBlank ("t"/"f"/""). */
export function stringifyVmlTrueFalseBlank(value: VmlTrueFalseBlank): string {
  return value === "" ? "" : stringifyVmlTrueFalse(value);
}

/** Parse a boolean-or-empty ST_TrueFalseBlank. */
export function parseVmlTrueFalseBlank(value: string): VmlTrueFalseBlank {
  return value === "" ? "" : parseVmlTrueFalse(value);
}

/** Inverse of {@link stringifyVmlAttributes} — fills the tabulated fields onto `out`. */
export function parseVmlAttributes(
  el: XmlElement,
  specs: readonly VmlAttrSpec[],
  out: Record<string, unknown>,
): void {
  for (const spec of specs) {
    const raw = el.attributes?.[spec.attr];
    if (raw === undefined) continue;
    out[spec.field] =
      spec.kind === "trueFalse"
        ? parseVmlTrueFalse(String(raw))
        : spec.kind === "trueFalseBlank"
          ? parseVmlTrueFalseBlank(String(raw))
          : spec.kind === "number"
            ? Number(raw)
            : String(raw);
  }
}
