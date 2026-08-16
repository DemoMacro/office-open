/**
 * SmartArt quick-style definition (dgm:styleDef, CT_StyleDefinition) — the
 * authoring-layer model describing how a diagram's style labels are formatted.
 *
 * Bidirectional: stringify builds the quickStyle part body, parse reads one
 * back into the same options shape, so custom styles round-trip intact.
 *
 * Reference: ISO/IEC 29500-4, dml-diagram.xsd (CT_StyleDefinition,
 * CT_StyleLabel, CT_TextProps and the EG_Text3D group from dml-main.xsd).
 *
 * @module
 */

import { attr, attrs, findChild } from "@office-open/xml";
import { OOXML_XML_DECLARATION } from "@office-open/xml";
import { stringify as stringifyInnerXml } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type { CustomDescriptor } from "../descriptor";
import type { ReadContext, WriteContext } from "../descriptor/context";
import type {
  DiagramCategoryOptions,
  DiagramDescriptionOptions,
  DiagramNameOptions,
} from "../drawing/diagram/headers";
import type { Scene3DOptions } from "../drawing/three-d/scene-3d";
import type { Shape3DOptions } from "../drawing/three-d/shape-3d";
import { scene3DDesc, shape3DDesc } from "../drawing/three-d/three-d-descriptors";
import { parseShapeStyle, stringifyShapeStyle } from "../theme/style-matrix";
import type { DefaultShapeStyleOptions } from "../theme/theme-options";
import { withDiagramNamespaces } from "./layout-definition";

/** dgm:txPr (CT_TextProps) — a choice of 3D text treatments (EG_Text3D). */
export interface TextProperties3DOptions {
  /** 3D shape properties for the label text (a:sp3d). */
  shape3d?: Shape3DOptions;
  /** Flatten text to a plane at this z coordinate, EMU (a:flatTx `@z`). */
  flatText?: number;
}

/** dgm:styleLbl (CT_StyleLabel) — formatting for one style slot. */
export interface StyleLabelOptions {
  /** Slot name referenced by dgm:layoutNode `@styleLbl` (required). */
  name: string;
  /** Shared 3D scene for this label (dgm:scene3d). */
  scene3d?: Scene3DOptions;
  /** 3D shape properties (dgm:sp3d). */
  shape3d?: Shape3DOptions;
  /** 3D text properties (dgm:txPr). */
  textProperties?: TextProperties3DOptions;
  /** Theme style-matrix references (dgm:style, a:CT_ShapeStyle). */
  style?: DefaultShapeStyleOptions;
  /** Raw a:extLst inner XML — verbatim round-trip. */
  ext?: string;
}

/** dgm:styleDef (CT_StyleDefinition) — the quickStyle part root. */
export interface StyleDefinitionOptions {
  /** Style identity URI (dgm:styleDef `@uniqueId`). */
  uniqueId?: string;
  /** Minimum Office version (dgm:styleDef `@minVer`). */
  minVer?: string;
  /** Localized titles (dgm:title*). */
  titles?: DiagramNameOptions[];
  /** Localized descriptions (dgm:desc*). */
  descriptions?: DiagramDescriptionOptions[];
  /** Gallery categories (dgm:catLst). */
  categories?: DiagramCategoryOptions[];
  /** Shared 3D scene for the whole style (dgm:scene3d). */
  scene3d?: Scene3DOptions;
  /** Style slots, at least one in a real part (dgm:styleLbl*). */
  styleLabels?: StyleLabelOptions[];
  /** Raw a:extLst inner XML — verbatim round-trip. */
  ext?: string;
}

// Noop contexts: style definitions hold theme references and plain colors,
// never media or relationships of their own.
const DIRECT_WRITE_CTX: WriteContext = {
  addRelationship: () => "",
  addMedia: () => "",
  addHyperlink: () => {},
};

const DIRECT_READ_CTX: ReadContext = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
};

// ── Stringify ──

// CT_StyleLabel declares scene3d/sp3d/style as local dgm-namespace elements
// wrapping dml-main content, so only the outer tag switches prefix; the
// dml-main children (a:camera, a:lnRef, …) keep theirs.
function renamespaceTag(tag: string, xml: string): string {
  return xml.replace(`<a:${tag}`, `<dgm:${tag}`).replace(`</a:${tag}>`, `</dgm:${tag}>`);
}

function stringifyTextProperties(o: TextProperties3DOptions, ctx: WriteContext): string {
  const body = o.shape3d
    ? (shape3DDesc.stringify(o.shape3d, ctx) ?? "")
    : o.flatText !== undefined
      ? `<a:flatTx${attrs({ z: o.flatText })}/>`
      : "";
  return body ? `<dgm:txPr>${body}</dgm:txPr>` : "";
}

function stringifyStyleLabel(o: StyleLabelOptions, ctx: WriteContext): string {
  let body = "";
  if (o.scene3d) body += renamespaceTag("scene3d", scene3DDesc.stringify(o.scene3d, ctx) ?? "");
  if (o.shape3d) body += renamespaceTag("sp3d", shape3DDesc.stringify(o.shape3d, ctx) ?? "");
  if (o.textProperties) body += stringifyTextProperties(o.textProperties, ctx);
  if (o.style) body += renamespaceTag("style", stringifyShapeStyle(o.style, ctx));
  if (o.ext) body += `<a:extLst>${o.ext}</a:extLst>`;
  return `<dgm:styleLbl${attrs({ name: o.name })}>${body}</dgm:styleLbl>`;
}

function stringifyLocalized(
  tag: string,
  o: DiagramNameOptions | DiagramDescriptionOptions,
): string {
  return `<dgm:${tag}${attrs({ lang: o.lang, val: o.val })}/>`;
}

/**
 * Serialize a style definition to the dgm:styleDef element (no XML
 * declaration, no namespace declarations — the part wrapper adds both).
 */
export function stringifyStyleDefinition(o: StyleDefinitionOptions): string {
  const body =
    (o.titles ?? []).map((t) => stringifyLocalized("title", t)).join("") +
    (o.descriptions ?? []).map((d) => stringifyLocalized("desc", d)).join("") +
    (o.categories?.length
      ? `<dgm:catLst>${o.categories
          .map((c) => `<dgm:cat${attrs({ type: c.type, pri: c.pri })}/>`)
          .join("")}</dgm:catLst>`
      : "") +
    (o.scene3d
      ? renamespaceTag("scene3d", scene3DDesc.stringify(o.scene3d, DIRECT_WRITE_CTX) ?? "")
      : "") +
    (o.styleLabels ?? []).map((l) => stringifyStyleLabel(l, DIRECT_WRITE_CTX)).join("") +
    (o.ext ? `<a:extLst>${o.ext}</a:extLst>` : "");
  return `<dgm:styleDef${attrs({ uniqueId: o.uniqueId, minVer: o.minVer })}>${body}</dgm:styleDef>`;
}

// ── Parse ──

function parseTextProperties(el: Element): TextProperties3DOptions {
  const result: Partial<TextProperties3DOptions> = {};
  const sp3d = findChild(el, "a:sp3d");
  if (sp3d) result.shape3d = shape3DDesc.parse(sp3d, DIRECT_READ_CTX);
  const flatTx = findChild(el, "a:flatTx");
  if (flatTx) {
    const z = attr(flatTx, "z");
    result.flatText = z !== undefined ? Number(z) : 0;
  }
  return result as TextProperties3DOptions;
}

function parseStyleLabel(el: Element): StyleLabelOptions {
  const result: Partial<StyleLabelOptions> = { name: attr(el, "name") ?? "" };
  const scene3d = findChild(el, "dgm:scene3d");
  if (scene3d) result.scene3d = scene3DDesc.parse(scene3d, DIRECT_READ_CTX);
  const sp3d = findChild(el, "dgm:sp3d");
  if (sp3d) result.shape3d = shape3DDesc.parse(sp3d, DIRECT_READ_CTX);
  const txPr = findChild(el, "dgm:txPr");
  if (txPr) result.textProperties = parseTextProperties(txPr);
  const style = findChild(el, "dgm:style");
  if (style) result.style = parseShapeStyle(style, DIRECT_READ_CTX);
  const extLst = findChild(el, "a:extLst");
  if (extLst) result.ext = stringifyInnerXml(extLst);
  return result as StyleLabelOptions;
}

function parseLocalized(
  el: Element,
  tag: string,
): DiagramNameOptions[] | DiagramDescriptionOptions[] | undefined {
  const out: (DiagramNameOptions | DiagramDescriptionOptions)[] = [];
  for (const child of el.elements ?? []) {
    if (child.name !== `dgm:${tag}`) continue;
    const val = attr(child, "val") ?? "";
    const lang = attr(child, "lang");
    out.push(lang ? { lang, val } : { val });
  }
  return out.length ? (out as DiagramNameOptions[]) : undefined;
}

/** Parse a dgm:styleDef element into options. */
export function parseStyleDefinition(el: Element): StyleDefinitionOptions {
  const root = el.name === "dgm:styleDef" ? el : (findChild(el, "dgm:styleDef") ?? el);
  const result: Partial<StyleDefinitionOptions> = {};
  const uniqueId = attr(root, "uniqueId");
  if (uniqueId) result.uniqueId = uniqueId;
  const minVer = attr(root, "minVer");
  if (minVer) result.minVer = minVer;
  const titles = parseLocalized(root, "title");
  if (titles) result.titles = titles;
  const descriptions = parseLocalized(root, "desc");
  if (descriptions) result.descriptions = descriptions;
  const catLst = findChild(root, "dgm:catLst");
  if (catLst) {
    const categories = (catLst.elements ?? [])
      .filter((c) => c.name === "dgm:cat")
      .map((c) => ({ type: attr(c, "type") ?? "", pri: Number(attr(c, "pri") ?? 0) }));
    if (categories.length) result.categories = categories;
  }
  const scene3d = findChild(root, "dgm:scene3d");
  if (scene3d) result.scene3d = scene3DDesc.parse(scene3d, DIRECT_READ_CTX);
  const styleLabels: StyleLabelOptions[] = [];
  for (const child of root.elements ?? []) {
    if (child.name === "dgm:styleLbl") styleLabels.push(parseStyleLabel(child));
  }
  if (styleLabels.length) result.styleLabels = styleLabels;
  const extLst = findChild(root, "a:extLst");
  if (extLst) result.ext = stringifyInnerXml(extLst);
  return result as StyleDefinitionOptions;
}

/** quickStyle part descriptor — stringify emits the element, parse reads it. */
export const styleDefDesc: CustomDescriptor<StyleDefinitionOptions> = {
  kind: "custom",
  stringify: (opts) => stringifyStyleDefinition(opts),
  parse: (el) => parseStyleDefinition(el),
};

/** Full quickStyle part body: XML declaration + namespaces + the element. */
export function stringifyStyleDefinitionPart(o: StyleDefinitionOptions): string {
  return OOXML_XML_DECLARATION + withDiagramNamespaces(stringifyStyleDefinition(o));
}
