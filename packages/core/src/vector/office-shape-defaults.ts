/**
 * Office shape defaults and layout — o:shapedefaults / o:shapelayout and
 * their children (colormru, colormenu, idmap, regrouptable, rules).
 *
 * These head the legacy VML drawing parts: docx settings.xml carries
 * shapedefaults/shapelayout (hdrShapeDefaults/shapeDefaults), and every
 * xlsx vmlDrawing part opens with `<o:shapelayout><o:idmap …/></o:shapelayout>`
 * plus the comment-shape `<o:shapedefaults v:ext="edit">` block.
 *
 * Reference: ISO/IEC 29500-4, vml-officeDrawing.xsd, CT_ShapeDefaults /
 * CT_ShapeLayout and their nested types.
 *
 * @module
 */
import type { Element as XmlElement } from "@office-open/xml";

import {
  stringifyVmlAttributes,
  parseVmlAttributes,
  type VmlAttrSpec,
  type VmlColor,
  type VmlTrueFalse,
} from "./attributes";
import { stringifyVmlFill, parseVmlFill, type VmlFillOptions } from "./shape-elements/fill";
import type { VmlExtAttribute } from "./shape-elements/office-elements";
import {
  stringifyVmlSkew,
  parseVmlSkew,
  stringifyVmlExtrusion,
  parseVmlExtrusion,
  stringifyVmlCallout,
  parseVmlCallout,
  stringifyVmlLock,
  parseVmlLock,
  type VmlSkewOptions,
  type VmlExtrusionOptions,
  type VmlCalloutOptions,
  type VmlLockOptions,
} from "./shape-elements/office-elements";
import { stringifyVmlShadow, parseVmlShadow, type VmlShadowOptions } from "./shape-elements/shadow";
import { stringifyVmlStroke, parseVmlStroke, type VmlStrokeOptions } from "./shape-elements/stroke";
import {
  stringifyVmlTextbox,
  parseVmlTextbox,
  type VmlTextboxOptions,
} from "./shape-elements/textbox";

// ── o:shapedefaults ──

/** o:shapedefaults options (CT_ShapeDefaults). */
export interface VmlShapeDefaultsOptions extends VmlExtAttribute {
  /** Highest shape id handed out so far. */
  spidmax?: number;
  style?: string;
  fill?: VmlTrueFalse;
  fillcolor?: VmlColor;
  stroke?: VmlTrueFalse;
  strokecolor?: VmlColor;
  /** form="qualified" — serialized as `o:allowincell`. */
  allowincell?: VmlTrueFalse;
  // xsd:all children, each at most once
  fillElement?: VmlFillOptions;
  strokeElement?: VmlStrokeOptions;
  textbox?: VmlTextboxOptions;
  shadow?: VmlShadowOptions;
  skew?: VmlSkewOptions;
  extrusion?: VmlExtrusionOptions;
  callout?: VmlCalloutOptions;
  lock?: VmlLockOptions;
  colormru?: VmlColorMruOptions;
  colormenu?: VmlColorMenuOptions;
}

const SHAPEDEFAULTS_ATTRS: readonly VmlAttrSpec[] = [
  { field: "ext", attr: "v:ext", kind: "string" },
  { field: "spidmax", attr: "spidmax", kind: "number" },
  { field: "style", attr: "style", kind: "string" },
  { field: "fill", attr: "fill", kind: "trueFalse" },
  { field: "fillcolor", attr: "fillcolor", kind: "string" },
  { field: "stroke", attr: "stroke", kind: "trueFalse" },
  { field: "strokecolor", attr: "strokecolor", kind: "string" },
  { field: "allowincell", attr: "o:allowincell", kind: "trueFalse" },
];

/** Serialize o:shapedefaults. */
export function stringifyVmlShapeDefaults(opts: VmlShapeDefaultsOptions): string {
  const children: string[] = [];
  if (opts.fillElement !== undefined) children.push(stringifyVmlFill(opts.fillElement));
  if (opts.strokeElement !== undefined) children.push(stringifyVmlStroke(opts.strokeElement));
  if (opts.textbox !== undefined) children.push(stringifyVmlTextbox(opts.textbox));
  if (opts.shadow !== undefined) children.push(stringifyVmlShadow(opts.shadow));
  if (opts.skew !== undefined) children.push(stringifyVmlSkew(opts.skew));
  if (opts.extrusion !== undefined) children.push(stringifyVmlExtrusion(opts.extrusion));
  if (opts.callout !== undefined) children.push(stringifyVmlCallout(opts.callout));
  if (opts.lock !== undefined) children.push(stringifyVmlLock(opts.lock));
  if (opts.colormru !== undefined) children.push(stringifyVmlColorMru(opts.colormru));
  if (opts.colormenu !== undefined) children.push(stringifyVmlColorMenu(opts.colormenu));
  const attrStr = stringifyVmlAttributes(opts as Record<string, unknown>, SHAPEDEFAULTS_ATTRS);
  return children.length > 0
    ? `<o:shapedefaults${attrStr}>${children.join("")}</o:shapedefaults>`
    : `<o:shapedefaults${attrStr}/>`;
}

/** Parse an o:shapedefaults element. */
export function parseVmlShapeDefaults(el: XmlElement): VmlShapeDefaultsOptions {
  const out: Record<string, unknown> = {};
  parseVmlAttributes(el, SHAPEDEFAULTS_ATTRS, out);
  for (const child of el.elements ?? []) {
    if (child.type !== "element") continue;
    switch (child.name) {
      case "v:fill":
        out.fillElement = parseVmlFill(child);
        break;
      case "v:stroke":
        out.strokeElement = parseVmlStroke(child);
        break;
      case "v:textbox":
        out.textbox = parseVmlTextbox(child);
        break;
      case "v:shadow":
        out.shadow = parseVmlShadow(child);
        break;
      case "o:skew":
        out.skew = parseVmlSkew(child);
        break;
      case "o:extrusion":
        out.extrusion = parseVmlExtrusion(child);
        break;
      case "o:callout":
        out.callout = parseVmlCallout(child);
        break;
      case "o:lock":
        out.lock = parseVmlLock(child);
        break;
      case "o:colormru":
        out.colormru = parseVmlColorMru(child);
        break;
      case "o:colormenu":
        out.colormenu = parseVmlColorMenu(child);
        break;
    }
  }
  return out as VmlShapeDefaultsOptions;
}

// ── o:colormru / o:colormenu ──

/** o:colormru options (CT_ColorMru) — the recently-used color list. */
export interface VmlColorMruOptions extends VmlExtAttribute {
  colors?: string;
}

/** Serialize o:colormru. */
export function stringifyVmlColorMru(opts: VmlColorMruOptions): string {
  return `<o:colormru${stringifyVmlAttributes(opts as Record<string, unknown>, [
    { field: "ext", attr: "v:ext", kind: "string" },
    { field: "colors", attr: "colors", kind: "string" },
  ])}/>`;
}

/** Parse an o:colormru element. */
export function parseVmlColorMru(el: XmlElement): VmlColorMruOptions {
  const out: Record<string, unknown> = {};
  parseVmlAttributes(
    el,
    [
      { field: "ext", attr: "v:ext", kind: "string" },
      { field: "colors", attr: "colors", kind: "string" },
    ],
    out,
  );
  return out as VmlColorMruOptions;
}

/** o:colormenu options (CT_ColorMenu) — the palette slots. */
export interface VmlColorMenuOptions extends VmlExtAttribute {
  strokecolor?: VmlColor;
  fillcolor?: VmlColor;
  shadowcolor?: VmlColor;
  extrusioncolor?: VmlColor;
}

const COLORMENU_ATTRS: readonly VmlAttrSpec[] = [
  { field: "ext", attr: "v:ext", kind: "string" },
  { field: "strokecolor", attr: "strokecolor", kind: "string" },
  { field: "fillcolor", attr: "fillcolor", kind: "string" },
  { field: "shadowcolor", attr: "shadowcolor", kind: "string" },
  { field: "extrusioncolor", attr: "extrusioncolor", kind: "string" },
];

/** Serialize o:colormenu. */
export function stringifyVmlColorMenu(opts: VmlColorMenuOptions): string {
  return `<o:colormenu${stringifyVmlAttributes(opts as Record<string, unknown>, COLORMENU_ATTRS)}/>`;
}

/** Parse an o:colormenu element. */
export function parseVmlColorMenu(el: XmlElement): VmlColorMenuOptions {
  const out: Record<string, unknown> = {};
  parseVmlAttributes(el, COLORMENU_ATTRS, out);
  return out as VmlColorMenuOptions;
}

// ── o:shapelayout ──

/** CT_Entry — one regroup-table row. */
export interface VmlRegroupEntryOptions {
  new?: number;
  old?: number;
}

const ENTRY_ATTRS: readonly VmlAttrSpec[] = [
  { field: "new", attr: "new", kind: "number" },
  { field: "old", attr: "old", kind: "number" },
];

/** o:regrouptable options (CT_RegroupTable). */
export interface VmlRegroupTableOptions extends VmlExtAttribute {
  entries?: VmlRegroupEntryOptions[];
}

/** ST_RType. */
export type VmlRuleType = "arc" | "callout" | "connector" | "align";

/** ST_How. */
export type VmlRuleHow = "top" | "middle" | "bottom" | "left" | "center" | "right";

/** CT_Proxy — one rule proxy. */
export interface VmlRuleProxyOptions {
  start?: boolean | "";
  end?: boolean | "";
  idref?: string;
  connectloc?: number;
}

/** CT_R — one layout rule. */
export interface VmlRuleOptions {
  id: string;
  type?: VmlRuleType;
  how?: VmlRuleHow;
  idref?: string;
  proxies?: VmlRuleProxyOptions[];
}

/** o:rules options (CT_Rules). */
export interface VmlRulesOptions extends VmlExtAttribute {
  rules?: VmlRuleOptions[];
}

/** o:idmap options (CT_IdMap). */
export interface VmlIdMapOptions extends VmlExtAttribute {
  /** Space-separated shape-id prefixes, e.g. "1 2 3". */
  data?: string;
}

/** o:shapelayout options (CT_ShapeLayout). */
export interface VmlShapeLayoutOptions extends VmlExtAttribute {
  idmap?: VmlIdMapOptions;
  regrouptable?: VmlRegroupTableOptions;
  rules?: VmlRulesOptions;
}

/** Serialize an o:idmap. */
function stringifyVmlIdMap(opts: VmlIdMapOptions): string {
  return `<o:idmap${stringifyVmlAttributes(opts as Record<string, unknown>, [
    { field: "ext", attr: "v:ext", kind: "string" },
    { field: "data", attr: "data", kind: "string" },
  ])}/>`;
}

/** Parse an o:idmap element. */
function parseVmlIdMap(el: XmlElement): VmlIdMapOptions {
  const out: Record<string, unknown> = {};
  parseVmlAttributes(
    el,
    [
      { field: "ext", attr: "v:ext", kind: "string" },
      { field: "data", attr: "data", kind: "string" },
    ],
    out,
  );
  return out as VmlIdMapOptions;
}

/** Serialize an o:regrouptable. */
function stringifyVmlRegroupTable(opts: VmlRegroupTableOptions): string {
  const entries = (opts.entries ?? [])
    .map(
      (entry) =>
        `<o:entry${stringifyVmlAttributes(entry as Record<string, unknown>, ENTRY_ATTRS)}/>`,
    )
    .join("");
  const attrStr = stringifyVmlAttributes(opts as Record<string, unknown>, [
    { field: "ext", attr: "v:ext", kind: "string" },
  ]);
  return `<o:regrouptable${attrStr}>${entries}</o:regrouptable>`;
}

/** Parse an o:regrouptable element. */
function parseVmlRegroupTable(el: XmlElement): VmlRegroupTableOptions {
  const out: Record<string, unknown> = {};
  parseVmlAttributes(el, [{ field: "ext", attr: "v:ext", kind: "string" }], out);
  const entries: VmlRegroupEntryOptions[] = [];
  for (const child of el.elements ?? []) {
    if (child.type === "element" && child.name === "o:entry") {
      const entry: Record<string, unknown> = {};
      parseVmlAttributes(child, ENTRY_ATTRS, entry);
      entries.push(entry as VmlRegroupEntryOptions);
    }
  }
  if (entries.length > 0) out.entries = entries;
  return out as VmlRegroupTableOptions;
}

/** Serialize a proxy child inside an o:r. */
function stringifyVmlRuleProxy(proxy: VmlRuleProxyOptions): string {
  const attrs: string[] = [];
  if (proxy.start !== undefined) {
    attrs.push(`start="${proxy.start === "" ? "" : proxy.start ? "t" : "f"}"`);
  }
  if (proxy.end !== undefined) {
    attrs.push(`end="${proxy.end === "" ? "" : proxy.end ? "t" : "f"}"`);
  }
  if (proxy.idref !== undefined) attrs.push(`idref="${proxy.idref}"`);
  if (proxy.connectloc !== undefined) attrs.push(`connectloc="${proxy.connectloc}"`);
  return `<o:proxy${attrs.length > 0 ? " " + attrs.join(" ") : ""}/>`;
}

/** Parse an o:proxy element. */
function parseVmlRuleProxy(el: XmlElement): VmlRuleProxyOptions {
  const out: VmlRuleProxyOptions = {};
  const attrs = el.attributes ?? {};
  if (attrs.start !== undefined) {
    const raw = String(attrs.start);
    out.start = raw === "" ? "" : raw === "t" || raw === "true" || raw === "1";
  }
  if (attrs.end !== undefined) {
    const raw = String(attrs.end);
    out.end = raw === "" ? "" : raw === "t" || raw === "true" || raw === "1";
  }
  if (attrs.idref !== undefined) out.idref = String(attrs.idref);
  if (attrs.connectloc !== undefined) out.connectloc = Number(attrs.connectloc);
  return out;
}

/** Serialize an o:r (hand-written — `id` is required, so no table-driven cast). */
function stringifyVmlRule(rule: VmlRuleOptions): string {
  const attrs: string[] = [`id="${rule.id}"`];
  if (rule.type !== undefined) attrs.push(`type="${rule.type}"`);
  if (rule.how !== undefined) attrs.push(`how="${rule.how}"`);
  if (rule.idref !== undefined) attrs.push(`idref="${rule.idref}"`);
  const proxies = (rule.proxies ?? []).map(stringifyVmlRuleProxy).join("");
  const attrStr = attrs.length > 0 ? " " + attrs.join(" ") : "";
  return `<o:r${attrStr}>${proxies}</o:r>`;
}

/** Parse an o:r element. */
function parseVmlRule(el: XmlElement): VmlRuleOptions {
  const attrs = el.attributes ?? {};
  const out: VmlRuleOptions = { id: String(attrs.id ?? "") };
  if (attrs.type !== undefined) out.type = String(attrs.type) as VmlRuleType;
  if (attrs.how !== undefined) out.how = String(attrs.how) as VmlRuleHow;
  if (attrs.idref !== undefined) out.idref = String(attrs.idref);
  const proxies: VmlRuleProxyOptions[] = [];
  for (const child of el.elements ?? []) {
    if (child.type === "element" && child.name === "o:proxy") {
      proxies.push(parseVmlRuleProxy(child));
    }
  }
  if (proxies.length > 0) out.proxies = proxies;
  return out;
}

/** Serialize an o:rules block. */
function stringifyVmlRulesBlock(opts: VmlRulesOptions): string {
  const rules = (opts.rules ?? []).map(stringifyVmlRule).join("");
  const attrStr = stringifyVmlAttributes(opts as Record<string, unknown>, [
    { field: "ext", attr: "v:ext", kind: "string" },
  ]);
  return `<o:rules${attrStr}>${rules}</o:rules>`;
}

/** Parse an o:rules element. */
function parseVmlRulesBlock(el: XmlElement): VmlRulesOptions {
  const out: Record<string, unknown> = {};
  parseVmlAttributes(el, [{ field: "ext", attr: "v:ext", kind: "string" }], out);
  const rules: VmlRuleOptions[] = [];
  for (const child of el.elements ?? []) {
    if (child.type === "element" && child.name === "o:r") {
      rules.push(parseVmlRule(child));
    }
  }
  if (rules.length > 0) out.rules = rules;
  return out as VmlRulesOptions;
}

/** Serialize o:shapelayout. */
export function stringifyVmlShapeLayout(opts: VmlShapeLayoutOptions): string {
  const children: string[] = [];
  if (opts.idmap !== undefined) children.push(stringifyVmlIdMap(opts.idmap));
  if (opts.regrouptable !== undefined) children.push(stringifyVmlRegroupTable(opts.regrouptable));
  if (opts.rules !== undefined) children.push(stringifyVmlRulesBlock(opts.rules));
  const attrStr = stringifyVmlAttributes(opts as Record<string, unknown>, [
    { field: "ext", attr: "v:ext", kind: "string" },
  ]);
  return children.length > 0
    ? `<o:shapelayout${attrStr}>${children.join("")}</o:shapelayout>`
    : `<o:shapelayout${attrStr}/>`;
}

/** Parse an o:shapelayout element. */
export function parseVmlShapeLayout(el: XmlElement): VmlShapeLayoutOptions {
  const out: Record<string, unknown> = {};
  parseVmlAttributes(el, [{ field: "ext", attr: "v:ext", kind: "string" }], out);
  for (const child of el.elements ?? []) {
    if (child.type !== "element") continue;
    if (child.name === "o:idmap") out.idmap = parseVmlIdMap(child);
    else if (child.name === "o:regrouptable") out.regrouptable = parseVmlRegroupTable(child);
    else if (child.name === "o:rules") out.rules = parseVmlRulesBlock(child);
  }
  return out as VmlShapeLayoutOptions;
}
