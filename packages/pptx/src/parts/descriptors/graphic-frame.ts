/**
 * GraphicFrame non-visual frame properties (p:cNvGraphicFramePr) — shared by
 * the five p:graphicFrame descriptors (table/chart/smartArt/ole/lockedCanvas).
 *
 * @module
 */
import type { GraphicFrameLockingOptions } from "@office-open/core";
import { extUriMatches, parseOnOff } from "@office-open/core";
import type { ReadContext } from "@office-open/core/descriptor";
import { graphicFrameLockingDesc } from "@office-open/core/drawing";
import { attr, findChild } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";
import type { PlaceholderOrientation, PlaceholderSize, PlaceholderType } from "@shared/shape/shape";

/**
 * Serialize p:cNvGraphicFramePr. Locking is tri-state:
 * undefined → the fresh default (MS Office's noGrp="1" emission); null → an
 * empty frame-props element; an object → the explicit lock flags.
 */
export function stringifyCnvGraphicFramePr(locking?: GraphicFrameLockingOptions | null): string {
  if (locking === null) return "<p:cNvGraphicFramePr/>";
  if (locking === undefined)
    return '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>';
  // An explicit object always emits the locks element — empty flags keep the
  // bare <a:graphicFrameLocks/> form the source may carry.
  const locks = graphicFrameLockingDesc.stringify(locking, undefined as never) ?? "";
  return `<p:cNvGraphicFramePr>${locks || "<a:graphicFrameLocks/>"}</p:cNvGraphicFramePr>`;
}

/**
 * Read the frame locking from a p:nvGraphicFramePr parent. Returns undefined
 * when the element is absent, null for an explicit empty cNvGraphicFramePr,
 * and the parsed flags otherwise.
 */
export function readGraphicFrameLocking(
  nvGraphicFramePr: XmlElement | undefined,
  ctx: ReadContext,
): GraphicFrameLockingOptions | null | undefined {
  const cNvGfp = nvGraphicFramePr ? findChild(nvGraphicFramePr, "p:cNvGraphicFramePr") : undefined;
  if (!cNvGfp) return undefined;
  const locks = findChild(cNvGfp, "a:graphicFrameLocks");
  if (!locks) return null;
  return graphicFrameLockingDesc.parse(locks, ctx);
}

/**
 * Placeholder reference carried in p:nvPr > p:ph — applies to every drawing
 * (shapes, pictures, and all five graphicFrame kinds).
 */
export interface NvPrPlaceholderOptions {
  /** CT_Placeholder `@type` — ST_PlaceholderType. */
  placeholder?: PlaceholderType;
  /** CT_Placeholder `@idx`. */
  placeholderIndex?: number;
  /** CT_Placeholder `@sz` — sizing hint (default "full"). */
  placeholderSize?: PlaceholderSize;
  /** CT_Placeholder `@orient` — orientation hint (default "horz"). */
  placeholderOrientation?: PlaceholderOrientation;
  hasCustomPrompt?: boolean;
  /** p:nvPr `@isPhoto`. */
  isPhoto?: boolean;
  /** p:nvPr `@userDrawn`. */
  userDrawn?: boolean;
  /**
   * p14:modId `@val` from the nvPr extension list — Office's collaboration
   * modification stamp, the sole known p:nvPr ext content.
   */
  modId?: string;
}

/** The p14:modId extension uri (CT_ApplicationNonVisualDrawingProps extLst). */
const MODID_EXT_URI = "{D42A27DB-BD31-4B8C-83A1-F6EECF244321}";

/** Serialize the p:nvPr content (placeholder reference, or photo/user-drawn attrs). */
export function stringifyNvPr(opts: NvPrPlaceholderOptions): string {
  const modIdExt = opts.modId
    ? `<p:extLst><p:ext uri="${MODID_EXT_URI}"><p14:modId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="${opts.modId}"/></p:ext></p:extLst>`
    : "";
  // p:ph @type is optional (XSD default "obj"), so a placeholder keyed only
  // by idx must still emit <p:ph/> without a type attribute.
  if (opts.placeholder || opts.placeholderIndex !== undefined) {
    const phAttrs: string[] = [];
    if (opts.placeholder) phAttrs.push(`type="${opts.placeholder}"`);
    if (opts.placeholderIndex !== undefined) phAttrs.push(`idx="${opts.placeholderIndex}"`);
    if (opts.placeholderSize !== undefined) phAttrs.push(`sz="${opts.placeholderSize}"`);
    if (opts.placeholderOrientation !== undefined)
      phAttrs.push(`orient="${opts.placeholderOrientation}"`);
    if (opts.hasCustomPrompt) phAttrs.push('hasCustomPrompt="1"');
    return `<p:nvPr><p:ph ${phAttrs.join(" ")}/>${modIdExt}</p:nvPr>`;
  }
  if (opts.isPhoto || opts.userDrawn) {
    const nvPrAttrs: string[] = [];
    if (opts.isPhoto) nvPrAttrs.push('isPhoto="1"');
    if (opts.userDrawn) nvPrAttrs.push('userDrawn="1"');
    return modIdExt
      ? `<p:nvPr ${nvPrAttrs.join(" ")}>${modIdExt}</p:nvPr>`
      : `<p:nvPr ${nvPrAttrs.join(" ")}/>`;
  }
  return modIdExt ? `<p:nvPr>${modIdExt}</p:nvPr>` : "<p:nvPr/>";
}

/** Read the p:nvPr placeholder reference into a result object. */
export function readNvPrPlaceholder(
  nvParent: XmlElement,
  result: {
    placeholder?: PlaceholderType;
    placeholderIndex?: number;
    placeholderSize?: PlaceholderSize;
    placeholderOrientation?: PlaceholderOrientation;
    hasCustomPrompt?: boolean;
    isPhoto?: boolean;
    userDrawn?: boolean;
    modId?: string;
  },
): void {
  const nvPr = findChild(nvParent, "p:nvPr");
  if (!nvPr) return;
  if (nvPr.attributes) {
    if (nvPr.attributes["isPhoto"] !== undefined)
      result.isPhoto = parseOnOff(nvPr.attributes["isPhoto"]) ?? false;
    if (nvPr.attributes["userDrawn"] !== undefined)
      result.userDrawn = parseOnOff(nvPr.attributes["userDrawn"]) ?? false;
  }
  const ph = findChild(nvPr, "p:ph");
  if (ph?.attributes) {
    if (ph.attributes["type"] !== undefined)
      result.placeholder = String(ph.attributes["type"]) as PlaceholderType;
    if (ph.attributes["idx"] !== undefined) result.placeholderIndex = Number(ph.attributes["idx"]);
    if (ph.attributes["sz"] !== undefined)
      result.placeholderSize = ph.attributes["sz"] as PlaceholderSize;
    if (ph.attributes["orient"] !== undefined)
      result.placeholderOrientation = ph.attributes["orient"] as PlaceholderOrientation;
    if (ph.attributes["hasCustomPrompt"] !== undefined)
      result.hasCustomPrompt = parseOnOff(ph.attributes["hasCustomPrompt"]) ?? false;
  }
  const extLst = findChild(nvPr, "p:extLst");
  if (extLst) {
    for (const ext of extLst.elements ?? []) {
      if (ext.name !== "p:ext" || !extUriMatches(attr(ext, "uri"), MODID_EXT_URI)) continue;
      const modId = findChild(ext, "p14:modId");
      const val = modId ? attr(modId, "val") : undefined;
      if (val) result.modId = val;
    }
  }
}
