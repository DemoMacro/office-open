/**
 * GraphicFrame non-visual frame properties (p:cNvGraphicFramePr) — shared by
 * the five p:graphicFrame descriptors (table/chart/smartArt/ole/lockedCanvas).
 *
 * @module
 */
import type { GraphicFrameLockingOptions } from "@office-open/core";
import { extUriMatches, parseOnOff, xsdPlaceholderType } from "@office-open/core";
import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import {
  buildHyperlinkElement,
  graphicFrameLockingDesc,
  readHyperlink,
  registerHyperlink,
  stringifyNonVisualDrawingProperties,
} from "@office-open/core/drawing";
import type { TextHyperlinkOptions } from "@office-open/core/drawing";
import type { NonVisualDrawingPropertiesOptions } from "@office-open/core/drawing";
import { attr, findChild, stringifyElement } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";
import { parseCustDataLst, stringifyCustDataLst } from "@parts/slide/c-sld";
import type { SlideCustomerDataReferenceOptions } from "@shared/customer-data";
import type {
  NvPrExtensionOptions,
  PlaceholderOrientation,
  PlaceholderSize,
  PlaceholderType,
} from "@shared/shape/shape";

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
  placeholder?: PlaceholderType | true;
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
   * modification stamp.
   */
  modId?: string;
  /** Unknown application extensions (`p:nvPr/p:extLst`); `[]` preserves an empty list. */
  nvPrExtensions?: NvPrExtensionOptions[];
  /** Click hyperlink (a:hlinkClick inside p:cNvPr). */
  hyperlink?: TextHyperlinkOptions;
}

/** Graphic-frame nvPr adds the p:custDataLst slot shared by its five frames. */
export interface GraphicFrameNvPrOptions extends NvPrPlaceholderOptions {
  /** Customer-data references (p:nvPr > p:custDataLst > p:tags/p:custData). */
  customerData?: SlideCustomerDataReferenceOptions[];
}

export function stringifyGraphicFrameCnvPr(
  id: number | string,
  opts: NvPrPlaceholderOptions & Partial<NonVisualDrawingPropertiesOptions>,
  fallbackName: string,
  ctx: WriteContext,
): string {
  const key = opts.hyperlink ? registerHyperlink(opts.hyperlink, ctx) : undefined;
  const hyperlinkXml = opts.hyperlink
    ? buildHyperlinkElement("a:hlinkClick", opts.hyperlink, key)
    : undefined;
  return stringifyNonVisualDrawingProperties("p:cNvPr", id, opts, fallbackName, hyperlinkXml);
}

export function readGraphicFrameHyperlink(
  nvGraphicFramePr: XmlElement | undefined,
  ctx: ReadContext,
): TextHyperlinkOptions | undefined {
  const cNvPr = nvGraphicFramePr ? findChild(nvGraphicFramePr, "p:cNvPr") : undefined;
  const hlinkClick = cNvPr ? findChild(cNvPr, "a:hlinkClick") : undefined;
  return hlinkClick ? readHyperlink(hlinkClick, ctx) : undefined;
}

/** The p14:modId extension uri (CT_ApplicationNonVisualDrawingProps extLst). */
const MODID_EXT_URI = "{D42A27DB-BD31-4B8C-83A1-F6EECF244321}";

/** Serialize the p:nvPr content (placeholder reference, or photo/user-drawn attrs). */
export function stringifyNvPr(opts: GraphicFrameNvPrOptions): string {
  const customerDataXml = stringifyCustDataLst(opts.customerData);
  const extensions: NvPrExtensionOptions[] = [];
  if (opts.modId) {
    extensions.push({
      uri: MODID_EXT_URI,
      content:
        '<p14:modId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" ' +
        `val="${opts.modId}"/>`,
    });
  }
  extensions.push(...(opts.nvPrExtensions ?? []));
  const nvPrExtXml =
    opts.nvPrExtensions !== undefined || extensions.length > 0
      ? `<p:extLst>${extensions
          .map((extension) => `<p:ext uri="${extension.uri}">${extension.content ?? ""}</p:ext>`)
          .join("")}</p:extLst>`
      : "";
  // isPhoto/userDrawn live on p:nvPr itself, orthogonal to the p:ph child
  // (CT_ApplicationNonVisualDrawingProps) — a placeholder can carry both.
  const nvPrAttrs: string[] = [];
  if (opts.isPhoto) nvPrAttrs.push('isPhoto="1"');
  if (opts.userDrawn) nvPrAttrs.push('userDrawn="1"');
  const nvPrAttrsXml = nvPrAttrs.length ? ` ${nvPrAttrs.join(" ")}` : "";
  // p:ph @type is optional (XSD default "obj"), so a placeholder keyed only
  // by idx must still emit <p:ph/> without a type attribute.
  if (opts.placeholder || opts.placeholderIndex !== undefined) {
    const phAttrs: string[] = [];
    if (typeof opts.placeholder === "string")
      phAttrs.push(`type="${xsdPlaceholderType.to(opts.placeholder)}"`);
    if (opts.placeholderIndex !== undefined) phAttrs.push(`idx="${opts.placeholderIndex}"`);
    if (opts.placeholderSize !== undefined) phAttrs.push(`sz="${opts.placeholderSize}"`);
    if (opts.placeholderOrientation !== undefined)
      phAttrs.push(`orient="${opts.placeholderOrientation}"`);
    if (opts.hasCustomPrompt) phAttrs.push('hasCustomPrompt="1"');
    const phAttrsXml = phAttrs.length > 0 ? ` ${phAttrs.join(" ")}` : "";
    return `<p:nvPr${nvPrAttrsXml}><p:ph${phAttrsXml}/>${customerDataXml}${nvPrExtXml}</p:nvPr>`;
  }
  if (customerDataXml || nvPrExtXml)
    return `<p:nvPr${nvPrAttrsXml}>${customerDataXml}${nvPrExtXml}</p:nvPr>`;
  return nvPrAttrsXml ? `<p:nvPr${nvPrAttrsXml}/>` : "<p:nvPr/>";
}

/** Read the p:nvPr placeholder reference into a result object. */
export function readNvPrPlaceholder(
  nvParent: XmlElement,
  result: Partial<GraphicFrameNvPrOptions>,
): void {
  const nvPr = findChild(nvParent, "p:nvPr");
  if (!nvPr) return;
  if (nvPr.attributes) {
    if (nvPr.attributes["isPhoto"] !== undefined)
      result.isPhoto = parseOnOff(nvPr.attributes["isPhoto"]) ?? false;
    if (nvPr.attributes["userDrawn"] !== undefined)
      result.userDrawn = parseOnOff(nvPr.attributes["userDrawn"]) ?? false;
  }
  result.customerData = parseCustDataLst(findChild(nvPr, "p:custDataLst"));
  const ph = findChild(nvPr, "p:ph");
  if (ph?.attributes && Object.keys(ph.attributes).length > 0) {
    if (ph.attributes["type"] !== undefined)
      result.placeholder = xsdPlaceholderType.from(
        String(ph.attributes["type"]),
      ) as PlaceholderType;
    if (ph.attributes["idx"] !== undefined) result.placeholderIndex = Number(ph.attributes["idx"]);
    if (ph.attributes["sz"] !== undefined)
      result.placeholderSize = ph.attributes["sz"] as PlaceholderSize;
    if (ph.attributes["orient"] !== undefined)
      result.placeholderOrientation = ph.attributes["orient"] as PlaceholderOrientation;
    if (ph.attributes["hasCustomPrompt"] !== undefined)
      result.hasCustomPrompt = parseOnOff(ph.attributes["hasCustomPrompt"]) ?? false;
  } else if (ph) {
    result.placeholder = true;
  }
  const extLst = findChild(nvPr, "p:extLst");
  if (extLst) {
    const unknownExtensions = (extLst.elements ?? [])
      .filter((ext) => ext.name === "p:ext" && !extUriMatches(attr(ext, "uri"), MODID_EXT_URI))
      .map((ext) => ({
        uri: attr(ext, "uri") ?? "",
        ...(ext.elements?.length
          ? { content: ext.elements.map((child) => stringifyElement(child)).join("") }
          : {}),
      }));
    result.nvPrExtensions = unknownExtensions;
    for (const ext of extLst.elements ?? []) {
      if (ext.name !== "p:ext" || !extUriMatches(attr(ext, "uri"), MODID_EXT_URI)) continue;
      const modId = findChild(ext, "p14:modId");
      const val = modId ? attr(modId, "val") : undefined;
      if (val) result.modId = val;
    }
  }
}
