/**
 * Format scheme / style matrix (a:fmtScheme / CT_StyleMatrix) and shape style
 * (a:style / CT_ShapeStyle) stringify + parse.
 *
 * Reuses core fill/outline/effect/3d descriptors so the style matrix round-trips
 * through the same machinery as shape properties.
 *
 * @module
 */
import { findChild } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";

import type { ReadContext, WriteContext } from "../descriptor";
import {
  effectListDesc,
  fillDesc,
  outlineDesc,
  parseColorChoice,
  scene3DDesc,
  shape3DDesc,
  stringifyColorChoice,
} from "../drawingml";
import type { SolidFillOptions } from "../drawingml";
import type {
  EffectStyleOptions,
  FontReferenceOptions,
  FormatSchemeOptions,
  DefaultShapeStyleOptions,
  StyleMatrixReferenceOptions,
} from "./theme-options";

// ── Style matrix reference (a:lnRef / a:fillRef / a:effectRef) ──

function stringifyStyleMatrixReference(
  tag: string,
  opts: StyleMatrixReferenceOptions,
  ctx: WriteContext,
): string {
  const color = opts.color ? stringifyColorChoice(opts.color, ctx) : "";
  return `<${tag} idx="${opts.index}">${color}</${tag}>`;
}

function parseStyleMatrixReference(
  el: XmlElement | undefined,
  ctx: ReadContext,
): StyleMatrixReferenceOptions | undefined {
  if (!el?.attributes) return undefined;
  const idx = el.attributes["idx"];
  if (idx === undefined) return undefined;
  const parsedColor = parseColorChoice(el, ctx);
  const color = parsedColor && Object.keys(parsedColor).length > 0 ? parsedColor : undefined;
  return color ? { index: Number(idx), color } : { index: Number(idx) };
}

// ── Font reference (a:fontRef) ──

function stringifyFontReference(opts: FontReferenceOptions, ctx: WriteContext): string {
  const color = opts.color ? stringifyColorChoice(opts.color, ctx) : "";
  return `<a:fontRef idx="${opts.collection}">${color}</a:fontRef>`;
}

function parseFontReference(
  el: XmlElement | undefined,
  ctx: ReadContext,
): FontReferenceOptions | undefined {
  if (!el?.attributes) return undefined;
  const idx = el.attributes["idx"];
  if (idx !== "major" && idx !== "minor" && idx !== "none") return undefined;
  const parsedColor = parseColorChoice(el, ctx);
  const color: SolidFillOptions | undefined =
    parsedColor && Object.keys(parsedColor).length > 0 ? parsedColor : undefined;
  return color ? { collection: idx, color } : { collection: idx };
}

// ── Shape style (a:style / CT_ShapeStyle) ──

export function stringifyShapeStyle(opts: DefaultShapeStyleOptions, ctx: WriteContext): string {
  const lnRef = stringifyStyleMatrixReference("a:lnRef", opts.lineReference, ctx);
  const fillRef = stringifyStyleMatrixReference("a:fillRef", opts.fillReference, ctx);
  const effectRef = stringifyStyleMatrixReference("a:effectRef", opts.effectReference, ctx);
  const fontRef = stringifyFontReference(opts.fontReference, ctx);
  return `<a:style>${lnRef}${fillRef}${effectRef}${fontRef}</a:style>`;
}

export function parseShapeStyle(
  el: XmlElement | undefined,
  ctx: ReadContext,
): DefaultShapeStyleOptions | undefined {
  if (!el) return undefined;
  const lineReference = parseStyleMatrixReference(findChild(el, "a:lnRef"), ctx);
  const fillReference = parseStyleMatrixReference(findChild(el, "a:fillRef"), ctx);
  const effectReference = parseStyleMatrixReference(findChild(el, "a:effectRef"), ctx);
  const fontReference = parseFontReference(findChild(el, "a:fontRef"), ctx);
  if (!lineReference || !fillReference || !effectReference || !fontReference) return undefined;
  return { lineReference, fillReference, effectReference, fontReference };
}

// ── Effect style (a:effectStyle / CT_EffectStyleItem) ──

function stringifyEffectStyle(opts: EffectStyleOptions, ctx: WriteContext): string {
  const effectLst = opts.effects
    ? (effectListDesc.stringify(opts.effects, ctx) ?? "")
    : "<a:effectLst/>";
  const scene3d = opts.scene3d ? (scene3DDesc.stringify(opts.scene3d, ctx) ?? "") : "";
  const shape3d = opts.shape3d ? (shape3DDesc.stringify(opts.shape3d, ctx) ?? "") : "";
  return `<a:effectStyle>${effectLst}${scene3d}${shape3d}</a:effectStyle>`;
}

function parseEffectStyle(
  el: XmlElement | undefined,
  ctx: ReadContext,
): EffectStyleOptions | undefined {
  if (!el) return undefined;
  const result: Partial<EffectStyleOptions> = {};
  const effectLst = findChild(el, "a:effectLst");
  if (effectLst) {
    const effects = effectListDesc.parse(effectLst, ctx);
    if (effects && Object.keys(effects).length > 0) result.effects = effects;
  }
  const scene3dEl = findChild(el, "a:scene3d");
  if (scene3dEl) result.scene3d = scene3DDesc.parse(scene3dEl, ctx);
  const shape3dEl = findChild(el, "a:sp3d");
  if (shape3dEl) result.shape3d = shape3DDesc.parse(shape3dEl, ctx);
  return result as EffectStyleOptions | undefined;
}

// ── Format scheme (a:fmtScheme / CT_StyleMatrix) ──

/** Serialize a:fmtScheme. */
export function stringifyFormatScheme(opts: FormatSchemeOptions, ctx: WriteContext): string {
  const fillStyles = opts.fillStyles.map((f) => fillDesc.stringify(f, ctx) ?? "").join("");
  const lineStyles = opts.lineStyles.map((l) => outlineDesc.stringify(l, ctx) ?? "").join("");
  const effectStyles = opts.effectStyles.map((e) => stringifyEffectStyle(e, ctx)).join("");
  const backgroundFillStyles = opts.backgroundFillStyles
    .map((f) => fillDesc.stringify(f, ctx) ?? "")
    .join("");
  const name = opts.name ?? "Office";
  return `<a:fmtScheme name="${name}"><a:fillStyleLst>${fillStyles}</a:fillStyleLst><a:lnStyleLst>${lineStyles}</a:lnStyleLst><a:effectStyleLst>${effectStyles}</a:effectStyleLst><a:bgFillStyleLst>${backgroundFillStyles}</a:bgFillStyleLst></a:fmtScheme>`;
}

/** Parse a:fmtScheme. */
export function parseFormatScheme(
  el: XmlElement | undefined,
  ctx: ReadContext,
): FormatSchemeOptions | undefined {
  if (!el) return undefined;
  const result: FormatSchemeOptions = {
    fillStyles: [],
    lineStyles: [],
    effectStyles: [],
    backgroundFillStyles: [],
  };
  const name = el.attributes?.["name"];
  if (name) result.name = String(name);

  const fillStyleLst = findChild(el, "a:fillStyleLst");
  if (fillStyleLst) {
    for (const child of fillStyleLst.elements ?? []) {
      const fill = fillDesc.parse(child, ctx);
      if (fill) result.fillStyles.push(fill);
    }
  }
  const lnStyleLst = findChild(el, "a:lnStyleLst");
  if (lnStyleLst) {
    for (const child of lnStyleLst.elements ?? []) {
      if (child.name !== "a:ln") continue;
      const line = outlineDesc.parse(child, ctx);
      if (line) result.lineStyles.push(line);
    }
  }
  const effectStyleLst = findChild(el, "a:effectStyleLst");
  if (effectStyleLst) {
    for (const child of effectStyleLst.elements ?? []) {
      if (child.name !== "a:effectStyle") continue;
      const effect = parseEffectStyle(child, ctx);
      if (effect) result.effectStyles.push(effect);
    }
  }
  const bgFillStyleLst = findChild(el, "a:bgFillStyleLst");
  if (bgFillStyleLst) {
    for (const child of bgFillStyleLst.elements ?? []) {
      const fill = fillDesc.parse(child, ctx);
      if (fill) result.backgroundFillStyles.push(fill);
    }
  }
  return result;
}
