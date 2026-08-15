/**
 * SmartArt color-transform definition (dgm:colorsDef, CT_ColorTransform) —
 * the authoring-layer model mapping style labels to color lists.
 *
 * Bidirectional: stringify builds the colors part body, parse reads one back
 * into the same options shape, so custom color transforms round-trip intact.
 *
 * Reference: ISO/IEC 29500-4, dml-diagram.xsd (CT_ColorTransform,
 * CT_CTStyleLabel, CT_Colors with its EG_ColorChoice children).
 *
 * @module
 */

import { attr, attrs, findChild } from "@office-open/xml";
import { OOXML_XML_DECLARATION } from "@office-open/xml";
import { stringify as stringifyInnerXml } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type { CustomDescriptor } from "../descriptor";
import { createColorList, parseColorList } from "../drawing/diagram/diagram-style";
import type { ColorListOptions, ColorStyleLabelOptions } from "../drawing/diagram/diagram-style";
import type {
  DiagramCategoryOptions,
  DiagramDescriptionOptions,
  DiagramNameOptions,
} from "../drawing/diagram/headers";
import { withDiagramNamespaces } from "./layout-definition";

/** dgm:colorsDef (CT_ColorTransform) — the colors part root. */
export interface ColorDefinitionOptions {
  /** Color-transform identity URI (dgm:colorsDef @uniqueId). */
  uniqueId?: string;
  /** Minimum Office version (dgm:colorsDef @minVer). */
  minVer?: string;
  /** Localized titles (dgm:title*). */
  titles?: DiagramNameOptions[];
  /** Localized descriptions (dgm:desc*). */
  descriptions?: DiagramDescriptionOptions[];
  /** Gallery categories (dgm:catLst). */
  categories?: DiagramCategoryOptions[];
  /** Color slots, at least one in a real part (dgm:styleLbl*). */
  styleLabels?: ColorStyleLabelOptions[];
  /** Raw a:extLst inner XML — verbatim round-trip. */
  ext?: string;
}

// CT_CTStyleLabel color lists in schema order: fill, line, effect, text line,
// text fill, text effect.
const COLOR_LIST_FIELDS: readonly (readonly [string, keyof ColorStyleLabelOptions])[] = [
  ["fillClrLst", "fillColorList"],
  ["linClrLst", "lineColorList"],
  ["effectClrLst", "effectColorList"],
  ["txLinClrLst", "textLineColorList"],
  ["txFillClrLst", "textFillColorList"],
  ["txEffectClrLst", "textEffectColorList"],
];

// ── Stringify ──

function stringifyLocalized(
  tag: string,
  o: DiagramNameOptions | DiagramDescriptionOptions,
): string {
  return `<dgm:${tag}${attrs({ lang: o.lang, val: o.val })}/>`;
}

function stringifyColorStyleLabel(o: ColorStyleLabelOptions): string {
  let body = "";
  for (const [tag, field] of COLOR_LIST_FIELDS) {
    const list = o[field] as ColorListOptions | undefined;
    if (list) body += createColorList(`dgm:${tag}`, list);
  }
  return `<dgm:styleLbl${attrs({ name: o.name })}>${body}</dgm:styleLbl>`;
}

/**
 * Serialize a color-transform definition to the dgm:colorsDef element (no XML
 * declaration, no namespace declarations — the part wrapper adds both).
 */
export function stringifyColorDefinition(o: ColorDefinitionOptions): string {
  const body =
    (o.titles ?? []).map((t) => stringifyLocalized("title", t)).join("") +
    (o.descriptions ?? []).map((d) => stringifyLocalized("desc", d)).join("") +
    (o.categories?.length
      ? `<dgm:catLst>${o.categories
          .map((c) => `<dgm:cat${attrs({ type: c.type, pri: c.pri })}/>`)
          .join("")}</dgm:catLst>`
      : "") +
    (o.styleLabels ?? []).map(stringifyColorStyleLabel).join("") +
    (o.ext ? `<a:extLst>${o.ext}</a:extLst>` : "");
  return `<dgm:colorsDef${attrs({ uniqueId: o.uniqueId, minVer: o.minVer })}>${body}</dgm:colorsDef>`;
}

// ── Parse ──

function parseColorStyleLabel(el: Element): ColorStyleLabelOptions {
  const result: Partial<ColorStyleLabelOptions> = { name: attr(el, "name") ?? "" };
  for (const [tag, field] of COLOR_LIST_FIELDS) {
    const listEl = findChild(el, `dgm:${tag}`);
    if (listEl) (result as Record<string, unknown>)[field] = parseColorList(listEl);
  }
  return result as ColorStyleLabelOptions;
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

/** Parse a dgm:colorsDef element into options. */
export function parseColorDefinition(el: Element): ColorDefinitionOptions {
  const root = el.name === "dgm:colorsDef" ? el : (findChild(el, "dgm:colorsDef") ?? el);
  const result: Partial<ColorDefinitionOptions> = {};
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
  const styleLabels: ColorStyleLabelOptions[] = [];
  for (const child of root.elements ?? []) {
    if (child.name === "dgm:styleLbl") styleLabels.push(parseColorStyleLabel(child));
  }
  if (styleLabels.length) result.styleLabels = styleLabels;
  const extLst = findChild(root, "a:extLst");
  if (extLst) result.ext = stringifyInnerXml(extLst);
  return result as ColorDefinitionOptions;
}

/** colors part descriptor — stringify emits the element, parse reads it. */
export const colorsDefDesc: CustomDescriptor<ColorDefinitionOptions> = {
  kind: "custom",
  stringify: (opts) => stringifyColorDefinition(opts),
  parse: (el) => parseColorDefinition(el),
};

/** Full colors part body: XML declaration + namespaces + the element. */
export function stringifyColorDefinitionPart(o: ColorDefinitionOptions): string {
  return OOXML_XML_DECLARATION + withDiagramNamespaces(stringifyColorDefinition(o));
}
