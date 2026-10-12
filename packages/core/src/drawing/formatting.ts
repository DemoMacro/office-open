/**
 * Shared DrawingML background and whole-shape formatting containers.
 *
 * Reference: ISO/IEC 29500-1, dml-main.xsd, CT_BackgroundFormatting and
 * CT_WholeE2oFormatting.
 *
 * @module
 */

import { findChild, stringifyElement } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import { parse } from "../descriptor";
import type { ReadContext, WriteContext } from "../descriptor/context";
import type { EffectDagOptions } from "./effects/effect-dag";
import { createEffectDag } from "./effects/effect-dag";
import { effectListDesc } from "./effects/effect-descriptors";
import type { EffectListOptions } from "./effects/effect-list";
import { fillDesc } from "./fill/fill-descriptors";
import type { FillOptions } from "./fill/fill-options";
import type { OutlineOptions } from "./outline/outline";
import { outlineDesc } from "./outline/outline-descriptors";

/** An OfficeArt extension (a:ext). Content may only contain foreign namespaces. */
export interface OfficeArtExtensionOptions {
  uri: string;
  content?: string;
}

/** Background formatting (CT_BackgroundFormatting). */
export interface BackgroundFormattingOptions {
  fill?: FillOptions | null;
  effects?: EffectListOptions;
  effectDag?: EffectDagOptions;
  extensions?: OfficeArtExtensionOptions[];
}

/** Whole-shape formatting (CT_WholeE2oFormatting). */
export interface WholeE2oFormattingOptions {
  outline?: OutlineOptions;
  effects?: EffectListOptions;
  effectDag?: EffectDagOptions;
  extensions?: OfficeArtExtensionOptions[];
}

function stringifyEffectChoice(
  options: { effects?: EffectListOptions; effectDag?: EffectDagOptions },
  ctx: WriteContext,
): string {
  if (options.effectDag) return createEffectDag(options.effectDag);
  if (options.effects) return effectListDesc.stringify(options.effects, ctx) ?? "";
  return "";
}

function stringifyExtensions(extensions: OfficeArtExtensionOptions[] | undefined): string {
  if (!extensions?.length) return "";
  const children = extensions
    .map((extension) => `<a:ext uri="${extension.uri}">${extension.content ?? ""}</a:ext>`)
    .join("");
  return `<a:extLst>${children}</a:extLst>`;
}

export function stringifyBackgroundFormatting(
  options: BackgroundFormattingOptions,
  ctx: WriteContext,
  containerPrefix = "a",
): string {
  const fill = options.fill == null ? "" : (fillDesc.stringify(options.fill, ctx) ?? "");
  const effects = stringifyEffectChoice(options, ctx);
  const extensions = stringifyExtensions(options.extensions);
  return `<${containerPrefix}:bg>${fill}${effects}${extensions}</${containerPrefix}:bg>`;
}

export function stringifyWholeE2oFormatting(
  options: WholeE2oFormattingOptions,
  ctx: WriteContext,
  containerPrefix = "a",
): string {
  const outline = options.outline ? (outlineDesc.stringify(options.outline, ctx) ?? "") : "";
  const effects = stringifyEffectChoice(options, ctx);
  const extensions = stringifyExtensions(options.extensions);
  return `<${containerPrefix}:whole>${outline}${effects}${extensions}</${containerPrefix}:whole>`;
}

function parseExtensions(parent: Element, prefix: string): OfficeArtExtensionOptions[] | undefined {
  const extensionList = findChild(parent, `${prefix}:extLst`);
  if (!extensionList) return undefined;
  const extensions = (extensionList.elements ?? [])
    .filter((child) => child.name === "a:ext")
    .map((extension) => ({
      uri: String(extension.attributes?.["uri"] ?? ""),
      ...(extension.elements?.length
        ? { content: (extension.elements ?? []).map(stringifyElement).join("") }
        : {}),
    }));
  return extensions.length > 0 ? extensions : [];
}

function parseEffectChoice(parent: Element): {
  effects?: EffectListOptions;
  effectDag?: EffectDagOptions;
} {
  const effectDag = findChild(parent, "a:effectDag");
  if (effectDag) return { effectDag: { raw: stringifyElement(effectDag) } };
  const effectList = findChild(parent, "a:effectLst");
  return effectList ? { effects: effectListDesc.parse(effectList, {} as ReadContext) } : {};
}

export function parseBackgroundFormatting(
  el: Element,
  ctx: ReadContext,
  containerPrefix = "a",
): BackgroundFormattingOptions {
  const result: BackgroundFormattingOptions = {};
  const fillChild =
    findChild(el, "a:noFill") ??
    findChild(el, "a:solidFill") ??
    findChild(el, "a:gradFill") ??
    findChild(el, "a:pattFill") ??
    findChild(el, "a:grpFill") ??
    findChild(el, "a:blipFill");
  if (fillChild) result.fill = parse(fillDesc, fillChild, ctx);
  Object.assign(result, parseEffectChoice(el));
  const extensions = parseExtensions(el, containerPrefix);
  if (extensions) result.extensions = extensions;
  return result;
}

export function parseWholeE2oFormatting(
  el: Element,
  ctx: ReadContext,
  containerPrefix = "a",
): WholeE2oFormattingOptions {
  const result: WholeE2oFormattingOptions = {};
  const outline = findChild(el, "a:ln");
  if (outline) result.outline = outlineDesc.parse(outline, ctx);
  Object.assign(result, parseEffectChoice(el));
  const extensions = parseExtensions(el, containerPrefix);
  if (extensions) result.extensions = extensions;
  return result;
}
