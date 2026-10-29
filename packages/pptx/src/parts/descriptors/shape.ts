/**
 * Shape (p:sp) and Picture (p:pic) descriptors for PPTX slides.
 *
 * These CustomDescriptor implementations produce the same XML output as the
 * class-based Shape and Picture components, but through the descriptor pipeline.
 *
 * @module
 */

import {
  parseOnOff,
  presetGeometryDesc,
  toUint8Array,
  xsdPlaceholderType,
} from "@office-open/core";
import type { NonVisualDrawingPropertiesOptions, ShapeLockingOptions } from "@office-open/core";
import type { CustomDescriptor, WriteContext, ReadContext } from "@office-open/core/descriptor";
import { parse, stringify } from "@office-open/core/descriptor";
import {
  shapeLockingDesc,
  effectListDesc,
  shapePropertiesDesc,
  scene3DDesc,
  shape3DDesc,
  textBodyDesc,
  stringifyNonVisualDrawingProperties,
  parseNonVisualDrawingProperties,
  pictureLockingDesc,
  blipDesc,
  customGeometryDesc,
  sourceRectangleDesc,
  stretchDesc,
  createSourceRectangle,
  fillDesc,
  findFillChild,
  outlineDesc,
  stringifyColorChoice,
  registerHyperlink,
  buildHyperlinkElement,
  readHyperlink,
} from "@office-open/core/drawing";
import type {
  BlipCompression,
  ShapeType,
  SourceRectangleOptions,
  Transform2DOptions,
} from "@office-open/core/drawing";
import type { Element as XmlElement } from "@office-open/xml";
import { findChild, findFirst, attrNum, attr, stringifyElement } from "@office-open/xml";
import { parseCustDataLst, stringifyCustDataLst } from "@parts/slide/c-sld";
import { imageTypeFromPath } from "@shared/media/image-type";
import type {
  AlternatePictureBlipFillOptions,
  AlternatePictureFillBranchOptions,
} from "@shared/picture";
import type { PictureOptions } from "@shared/picture";
import {
  readShapeStyle,
  type ShapeOptions,
  type ShapeStyleOptions,
  type StyleReferenceColor,
} from "@shared/shape/shape";

import type { PptxWriteContext, MediaEntry } from "../../context";
import { readNvPrPlaceholder, stringifyNvPr } from "./graphic-frame";
import { nextSlideDrawingId } from "./slide-drawing-ids";

// ── Auto-incrementing IDs ──

let _nextShapeId = 2;
let _nextPictureId = 100;

/** Reset shape ID counter (useful for tests). */
export function resetShapeIdCounter(value = 2): void {
  _nextShapeId = value;
}

// ── Shape (p:sp) descriptor ──

export const shapeDesc: CustomDescriptor<ShapeOptions> = {
  kind: "custom",

  stringify(opts, ctx) {
    const id = opts.id ?? nextSlideDrawingId(ctx) ?? _nextShapeId++;
    const name = opts.name ?? `Shape ${id}`;
    (ctx as PptxWriteContext).registerShapeId(name, id);
    const parts: string[] = [];

    // ── p:nvSpPr ──
    parts.push(stringifyNvSpPr(id, name, opts, ctx));

    // ── p:spPr ──
    const spPrXml = stringifySpPr(opts, ctx);
    if (spPrXml) parts.push(spPrXml);

    // ── p:style ──
    if (opts.style) {
      const styleXml = stringifyShapeStyle(opts.style, ctx);
      if (styleXml) parts.push(styleXml);
    }

    // ── p:txBody (optional in CT_Shape; omitted when absent so txBody-less
    // shapes like the notes sldImg placeholder round-trip without a spurious
    // empty body. Parse only sets textBody when a p:txBody element exists.)
    if (opts.textBody !== undefined) {
      parts.push(`<p:txBody>${textBodyDesc.stringify(opts.textBody, ctx) ?? ""}</p:txBody>`);
    }

    // ── Root attributes ──
    const spAttrs: string[] = [];
    if (opts.useBackgroundFill) spAttrs.push(' useBgFill="1"');

    return `<p:sp${spAttrs.join("")}>${parts.join("")}</p:sp>`;
  },

  parse(el, ctx) {
    const result: ShapeOptions = {};

    // Root attributes
    if (el.attributes) {
      if (el.attributes["useBgFill"] !== undefined)
        result.useBackgroundFill = parseOnOff(el.attributes["useBgFill"]) ?? false;
    }

    // p:nvSpPr — readNvSpPr returns the full ShapeOptions shape; assign whole
    // so fields added later flow without touching this whitelist.
    const nvSpPr = findChild(el, "p:nvSpPr");
    if (nvSpPr) {
      Object.assign(result, readNvSpPr(nvSpPr, ctx));
    }

    // p:spPr — transform goes top-level, paint fields land on `properties`;
    // @bwMode is a container attribute so it stays top-level too.
    const spPr = findChild(el, "p:spPr");
    if (spPr) {
      const { properties, ...transform } = readSpPr(spPr, ctx);
      Object.assign(result, transform);
      result.properties = properties;
      const bwMode = spPr.attributes?.["bwMode"];
      if (bwMode !== undefined) {
        result.blackWhiteMode = String(bwMode) as ShapeOptions["blackWhiteMode"];
      }
    }

    // p:style
    const style = findChild(el, "p:style");
    if (style) result.style = readShapeStyle(style, ctx);

    // p:txBody
    const txBody = findChild(el, "p:txBody");
    if (txBody) result.textBody = textBodyDesc.parse(txBody, ctx);

    return result;
  },
};

// ── Picture (p:pic) descriptor ──

export const pictureDesc: CustomDescriptor<PictureOptions> = {
  kind: "custom",

  stringify(opts, ctx) {
    const pptx = ctx as PptxWriteContext;
    const id = opts.id ?? nextSlideDrawingId(ctx) ?? _nextPictureId++;
    const name = opts.name ?? `Picture ${id}`;
    pptx.registerShapeId(name, id);

    // A linked-only picture (external URL, no bytes) registers no media —
    // its slide carries one External image relationship instead. A byteless
    // picture with no link either (a broken source reference) registers none
    // as well: there are no bytes to store, so the blip re-emits bare.
    let mediaFileName: string | undefined;
    if (opts.data !== undefined) {
      const data = toUint8Array(opts.data, { encoding: "base64" });
      mediaFileName = pptx.addMedia(data, opts.type, opts.fileName).slice(1, -1);
    }
    const blipReferenceId = opts.data !== undefined ? mediaFileName : opts.fileName;
    const blipRawReferenceId = opts.data !== undefined ? undefined : opts.relationshipId;

    const parts: string[] = [];

    // ── p:nvPicPr ──
    parts.push(stringifyNvPicPr(id, name, opts, ctx));

    // ── p:blipFill ──
    if (opts.alternateBlipFill) {
      parts.push(stringifyAlternateBlipFill(opts.alternateBlipFill, ctx));
    } else if (opts.blipFillPresent === "empty") {
      parts.push("<p:blipFill/>");
    } else if (opts.blipFillPresent !== false) {
      parts.push(
        stringifyPptxBlipFill(
          {
            fileName: blipReferenceId,
            rawReferenceId: blipRawReferenceId,
            sourceUrl: opts.sourceUrl,
            compression: opts.compression,
          },
          opts.sourceRectangle,
          opts.blipEffects,
          ctx,
          opts.useLocalDpi,
          opts.blipExt,
          opts.fillRectangle,
          opts.rotWithShape,
        ),
      );
    }

    // ── p:spPr ──
    const spPrXml = stringifyPicSpPr(opts, ctx);
    if (spPrXml) parts.push(spPrXml);

    // ── p:style ──
    if (opts.style) {
      const styleXml = stringifyShapeStyle(opts.style, ctx);
      if (styleXml) parts.push(styleXml);
    }

    return `<p:pic>${parts.join("")}</p:pic>`;
  },

  parse(el, ctx) {
    const result: Partial<PictureOptions> = {};

    // p:nvPicPr
    const nvPicPr = findChild(el, "p:nvPicPr");
    if (nvPicPr) {
      const cNvPr = findChild(nvPicPr, "p:cNvPr");
      Object.assign(result, parseNonVisualDrawingProperties(cNvPr));
      if (cNvPr?.attributes && cNvPr.attributes["id"] !== undefined) {
        result.id = Number(cNvPr.attributes["id"]);
      }
      const hlinkClick = cNvPr ? findChild(cNvPr, "a:hlinkClick") : undefined;
      if (hlinkClick) result.hyperlink = readHyperlink(hlinkClick, ctx);
      const cNvPicPr = findChild(nvPicPr, "p:cNvPicPr");
      if (cNvPicPr?.attributes?.preferRelativeResize !== undefined) {
        result.preferRelativeResize = cNvPicPr.attributes.preferRelativeResize === "1";
      }
      const picLocks = cNvPicPr ? findChild(cNvPicPr, "a:picLocks") : undefined;
      if (picLocks) {
        // Element presence is the signal — a bare <a:picLocks/> round-trips
        // as an empty locking object, not as a dropped element.
        result.locking = parse(pictureLockingDesc, picLocks, ctx) ?? {};
      }
      readNvPrPlaceholder(nvPicPr, result);
    }

    // p:spPr (position/size/flip/rotation)
    const spPr = findChild(el, "p:spPr");
    if (spPr) {
      const prstGeom = findChild(spPr, "a:prstGeom");
      // A source picture may carry no geometry at all — keep it suppressed.
      if (!prstGeom && !findChild(spPr, "a:custGeom")) {
        result.geometry = null;
      } else if (prstGeom) {
        const prst = attr(prstGeom, "prst");
        const avLst = findChild(prstGeom, "a:avLst");
        const hasGuides = (avLst?.elements ?? []).some((e) => e.name === "a:gd");
        if (hasGuides) {
          result.geometry = presetGeometryDesc.parse(prstGeom, ctx);
        } else if (prst && prst !== "rect") {
          // rect is the writer default — only a non-rect preset round-trips.
          result.geometry = prst as ShapeType;
        }
      }
      const customGeometry = findChild(spPr, "a:custGeom");
      if (customGeometry) result.customGeometry = customGeometryDesc.parse(customGeometry, ctx);
      const xfrm = findChild(spPr, "a:xfrm");
      if (xfrm) {
        const off = findChild(xfrm, "a:off");
        if (off?.attributes) {
          result.x = Number(off.attributes["x"] ?? 0);
          result.y = Number(off.attributes["y"] ?? 0);
        }
        const ext = findChild(xfrm, "a:ext");
        if (ext?.attributes) {
          result.width = Number(ext.attributes["cx"] ?? 0);
          result.height = Number(ext.attributes["cy"] ?? 0);
        }
        if (xfrm.attributes) {
          const flipH = parseOnOff(xfrm.attributes["flipH"]);
          if (flipH !== undefined) result.flipHorizontal = flipH;
          const flipV = parseOnOff(xfrm.attributes["flipV"]);
          if (flipV !== undefined) result.flipVertical = flipV;
          const rot = xfrm.attributes["rot"];
          if (rot !== undefined && rot !== "") result.rotation = Number(rot) / 60000;
        }
      }

      // Fill (a:noFill on cropped/locked pictures is common)
      const fillChild = findFillChild(spPr);
      if (fillChild) result.fill = parse(fillDesc, fillChild, ctx);

      // Outline (decorated pictures carry one)
      const picLn = findChild(spPr, "a:ln");
      if (picLn) result.outline = parse(outlineDesc, picLn, ctx);

      // Effects (e.g. shadow/reflection on the picture)
      const effectLst = findChild(spPr, "a:effectLst");
      if (effectLst) {
        const effects = parse(effectListDesc, effectLst, ctx);
        if (effects) result.effects = effects;
      }

      // 3D scene / shape properties
      const scene3d = findChild(spPr, "a:scene3d");
      if (scene3d) result.scene3d = scene3DDesc.parse(scene3d, ctx);
      const sp3d = findChild(spPr, "a:sp3d");
      if (sp3d) result.shape3d = shape3DDesc.parse(sp3d, ctx);
      const bwMode = spPr.attributes?.["bwMode"];
      if (bwMode !== undefined) result.blackWhiteMode = bwMode as PictureOptions["blackWhiteMode"];
      const spProperties = shapePropertiesDesc.parse(spPr, ctx);
      if (spProperties.extensions !== undefined) {
        result.shapePropertiesExtensions = spProperties.extensions;
      }
      if (spProperties.extensions !== undefined) {
        result.shapePropertiesExt = spProperties.ext ?? "";
      }
      if (spProperties.effectDag !== undefined) result.effectDag = spProperties.effectDag;
    }

    // p:style
    const picStyle = findChild(el, "p:style");
    if (picStyle) result.style = readShapeStyle(picStyle, ctx);

    // Crop rectangle from p:blipFill → a:srcRect (‰ → percent)
    const alternateBlipFill = findChild(el, "mc:AlternateContent");
    const blipFill = findChild(el, "p:blipFill");
    if (alternateBlipFill) result.alternateBlipFill = readAlternateBlipFill(alternateBlipFill, ctx);
    else if (!blipFill) result.blipFillPresent = false;
    else if (
      !findChild(blipFill, "a:blip") &&
      !findChild(blipFill, "a:srcRect") &&
      !findChild(blipFill, "a:stretch")
    )
      result.blipFillPresent = "empty";
    if (parseOnOff(blipFill?.attributes?.["rotWithShape"])) result.rotWithShape = true;
    const srcRectEl = blipFill ? findChild(blipFill, "a:srcRect") : undefined;
    if (srcRectEl) {
      const rect: SourceRectangleOptions = {};
      const l = attrNum(srcRectEl, "l");
      if (l !== undefined) rect.left = l / 1000;
      const t = attrNum(srcRectEl, "t");
      if (t !== undefined) rect.top = t / 1000;
      const r = attrNum(srcRectEl, "r");
      if (r !== undefined) rect.right = r / 1000;
      const b = attrNum(srcRectEl, "b");
      if (b !== undefined) rect.bottom = b / 1000;
      result.sourceRectangle = rect;
    }

    // Stretch fill rectangle — a bare <a:stretch/> (no a:fillRect child) is a
    // distinct emitted form some writers produce; keep it tri-state.
    const stretch = blipFill ? findChild(blipFill, "a:stretch") : undefined;
    if (stretch) {
      const fillRectEl = findChild(stretch, "a:fillRect");
      if (!fillRectEl) {
        result.fillRectangle = false;
      } else {
        const rect: SourceRectangleOptions = {};
        const l = attrNum(fillRectEl, "l");
        if (l !== undefined) rect.left = l / 1000;
        const t = attrNum(fillRectEl, "t");
        if (t !== undefined) rect.top = t / 1000;
        const r = attrNum(fillRectEl, "r");
        if (r !== undefined) rect.right = r / 1000;
        const b = attrNum(fillRectEl, "b");
        if (b !== undefined) rect.bottom = b / 1000;
        if (Object.keys(rect).length > 0) result.fillRectangle = rect;
      }
    }

    // Image data from p:blipFill → a:blip → r:embed / r:link
    const blip = findFirst(el, "a:blip");
    if (blip) {
      const parsedBlip = parse(blipDesc, blip, ctx);
      if (parsedBlip.referenceId !== undefined) result.relationshipId = parsedBlip.referenceId;
      if (parsedBlip.blipEffects) result.blipEffects = parsedBlip.blipEffects;
      if (parsedBlip.useLocalDpi !== undefined) result.useLocalDpi = parsedBlip.useLocalDpi;
      if (parsedBlip.ext !== undefined) result.blipExt = parsedBlip.ext;
      if (parsedBlip.compression !== undefined) result.compression = parsedBlip.compression;
      const rEmbed = attr(blip, "r:embed");
      if (rEmbed) {
        const imagePath = ctx.resolveRelationship(rEmbed);
        if (imagePath) {
          const imageData = ctx.getRaw(imagePath);
          if (imageData) {
            result.data = imageData;
            result.type = imageTypeFromPath(imagePath);
          }
          // Keep the source file name — re-deriving it from the shape name
          // renames the media part and forks it from its passthrough copy.
          result.fileName = imagePath.split("/").pop();
        }
      }
      const rLink = attr(blip, "r:link");
      if (rLink) {
        // External image source — the per-part rels keep External targets raw,
        // so resolveRelationship returns the URL itself.
        const url = ctx.resolveRelationship(rLink);
        if (url) {
          result.sourceUrl = url;
          if (!result.type) result.type = imageTypeFromPath(url);
        }
      }
    }

    // A picture whose rel target resolves to no part (a broken reference)
    // stays byteless: fabricating empty bytes would grow a 0-byte media part
    // the source never had. A linked-only picture is byteless by design.
    if (!result.type) result.type = "png";

    return result as PictureOptions;
  },
};

// ── Shape helper: p:nvSpPr ──

function stringifyNvSpPr(id: number, name: string, opts: ShapeOptions, ctx: WriteContext): string {
  // nvPr — @isPhoto/@userDrawn are the element's own attributes, orthogonal
  // to the p:ph child (CT_ApplicationNonVisualDrawingProps), so both emit
  // together. p:ph @type is optional (XSD default "obj"), so a placeholder
  // keyed only by idx must still emit <p:ph/> without a type attribute.
  const nvPrAttrs: string[] = [];
  if (opts.isPhoto) nvPrAttrs.push('isPhoto="1"');
  if (opts.userDrawn) nvPrAttrs.push('userDrawn="1"');
  const nvPrAttrsXml = nvPrAttrs.length > 0 ? ` ${nvPrAttrs.join(" ")}` : "";
  const customerDataXml = stringifyCustDataLst(opts.customerData);
  const nvPrExtensionXml =
    opts.nvPrExtensions !== undefined
      ? `<p:extLst>${opts.nvPrExtensions
          .map((extension) => `<p:ext uri="${extension.uri}">${extension.content ?? ""}</p:ext>`)
          .join("")}</p:extLst>`
      : "";
  const nvPrChildren: string[] = [];
  // The true marker preserves a source p:ph with no attributes.
  if (opts.placeholder || opts.placeholderIndex !== undefined) {
    const phAttrs: string[] = [];
    if (typeof opts.placeholder === "string")
      phAttrs.push(`type="${xsdPlaceholderType.to(opts.placeholder)}"`);
    if (opts.placeholderIndex !== undefined) phAttrs.push(`idx="${opts.placeholderIndex}"`);
    if (opts.placeholderSize !== undefined) phAttrs.push(`sz="${opts.placeholderSize}"`);
    if (opts.placeholderOrientation !== undefined)
      phAttrs.push(`orient="${opts.placeholderOrientation}"`);
    if (opts.hasCustomPrompt) phAttrs.push('hasCustomPrompt="1"');
    const phAttrsXml = phAttrs.join(" ");
    nvPrChildren.push(`<p:ph${phAttrsXml ? ` ${phAttrsXml}` : ""}/>`);
  }
  if (customerDataXml) nvPrChildren.push(customerDataXml);
  if (nvPrExtensionXml) nvPrChildren.push(nvPrExtensionXml);
  const nvPrContent =
    nvPrAttrs.length > 0 || nvPrChildren.length > 0
      ? `<p:nvPr${nvPrAttrsXml}>${nvPrChildren.join("")}</p:nvPr>`
      : "<p:nvPr/>";

  // cNvSpPr (with optional locking)
  const txBoxAttr = opts.textBox === undefined ? "" : ` txBox="${opts.textBox ? "1" : "0"}"`;
  let cNvSpPrContent = `<p:cNvSpPr${txBoxAttr}/>`;
  if (opts.locking) {
    const lockAttrs = buildLockAttrs(opts.locking);
    if (lockAttrs.length > 0) {
      cNvSpPrContent = `<p:cNvSpPr${txBoxAttr}><a:spLocks ${lockAttrs.join(" ")}/></p:cNvSpPr>`;
    } else {
      cNvSpPrContent = `<p:cNvSpPr${txBoxAttr}><a:spLocks/></p:cNvSpPr>`;
    }
  }

  // Click hyperlink (a:hlinkClick) lives inside p:cNvPr.
  let hlinkXml: string | undefined;
  if (opts.hyperlink) {
    const key = registerHyperlink(opts.hyperlink, ctx);
    hlinkXml = buildHyperlinkElement("a:hlinkClick", opts.hyperlink, key);
  }

  const cNvPrXml = stringifyNonVisualDrawingProperties("p:cNvPr", id, opts, name, hlinkXml);
  return `<p:nvSpPr>${cNvPrXml}${cNvSpPrContent}${nvPrContent}</p:nvSpPr>`;
}

function buildLockAttrs(opts: ShapeLockingOptions): string[] {
  const attrs: string[] = [];
  const keys = [
    "noGrp",
    "noSelect",
    "noRot",
    "noChangeAspect",
    "noMove",
    "noResize",
    "noEditPoints",
    "noAdjustHandles",
    "noChangeArrowheads",
    "noChangeShapeType",
    "noTextEdit",
  ] as const;
  for (const key of keys) {
    const val = opts[key];
    if (val !== undefined) attrs.push(`${key}="${val ? 1 : 0}"`);
  }
  return attrs;
}

// ── Shape helper: p:spPr ──

function stringifySpPr(opts: ShapeOptions, ctx: WriteContext): string {
  const pptx = ctx as PptxWriteContext;
  const sp = opts.properties ?? {};

  // Blip fill: pre-register via addImage so the deduped canonical keeps the
  // pinned source name on round-trip (imageType normalization like jpeg→jpg
  // would otherwise fork a second media part); fresh fills get the
  // pptx-local image_blip.<type>. shapePropertiesDesc's fillDesc re-registers
  // via addMedia, deduped to the same canonical.
  if (sp.fill && typeof sp.fill !== "string" && sp.fill.type === "blip" && sp.fill.data) {
    const blipFill = sp.fill;
    const raw = toUint8Array(blipFill.data!, { encoding: "base64" });
    const fileName = blipFill.fileName ?? `image_blip.${blipFill.imageType ?? "png"}`;
    pptx.addImage(fileName, {
      key: fileName,
      data: raw,
      fileName,
      type: (blipFill.imageType ?? "png") as MediaEntry["type"],
    });
  }

  // Placeholder shapes inherit geometry/fill from layout/master; only emit
  // them when explicitly set. Non-placeholder shapes default to rect geometry
  // and noFill. A p:ph with no @type (idx-only) is still a placeholder —
  // placeholderIndex alone carries its identity.
  const isPlaceholder = !!opts.placeholder || opts.placeholderIndex !== undefined;
  const geometry =
    sp.geometry === null
      ? undefined
      : sp.customGeometry
        ? undefined
        : sp.geometry === undefined
          ? isPlaceholder
            ? undefined
            : "rect"
          : sp.geometry;
  // A shape whose p:style carries a fillRef inherits its fill from the style
  // matrix — an explicit noFill would override that, so leave the spPr fill
  // unset unless the source (or the fresh caller) says otherwise. `fill: null`
  // marks a parsed source with no fill child — emit nothing there too.
  const styleOwnsFill = opts.style?.fillReference !== undefined;
  const fill =
    sp.fill === null
      ? undefined
      : isPlaceholder || styleOwnsFill
        ? sp.fill
        : (sp.fill ?? ({ type: "none" } as const));

  let spPrContent = shapePropertiesDesc.stringify(
    {
      x: opts.x,
      y: opts.y,
      width: opts.width,
      height: opts.height,
      flipHorizontal: opts.flipHorizontal,
      flipVertical: opts.flipVertical,
      rotation: opts.rotation,
      ...sp,
      geometry,
      fill,
    },
    ctx,
  );

  // @bwMode is a CT_ShapeProperties container attribute (CT_Shape itself only
  // has @useBgFill).
  const bwAttr = opts.blackWhiteMode ? ` bwMode="${opts.blackWhiteMode}"` : "";
  if (!spPrContent) return bwAttr ? `<p:spPr${bwAttr}/>` : "<p:spPr/>";
  return `<p:spPr${bwAttr}>${spPrContent}</p:spPr>`;
}

// ── Shape helper: p:style ──

/**
 * Serialize p:style (CT_ShapeStyle). Every ref carries a bare EG_ColorChoice
 * child (any of the six color kinds; fontRef included — @idx is
 * ST_FontCollectionIndex, not numeric). Shared by the shape descriptor and
 * the master placeholder emitter.
 */
export function stringifyShapeStyle(style: ShapeStyleOptions, ctx: WriteContext): string {
  const ref = (
    name: string,
    { index, color }: { index: number | string; color?: StyleReferenceColor },
  ): string => {
    if (color === undefined) return `<${name} idx="${index}"/>`;
    const choice = typeof color === "string" ? { value: color } : color;
    return `<${name} idx="${index}">${stringifyColorChoice(choice, ctx)}</${name}>`;
  };

  const parts: string[] = [];
  if (style.lineReference) parts.push(ref("a:lnRef", style.lineReference));
  if (style.fillReference) parts.push(ref("a:fillRef", style.fillReference));
  if (style.effectReference) parts.push(ref("a:effectRef", style.effectReference));
  if (style.fontReference)
    parts.push(
      ref("a:fontRef", {
        index: style.fontReference.collection,
        color: style.fontReference.color,
      }),
    );
  return parts.length > 0 ? `<p:style>${parts.join("")}</p:style>` : "";
}

// ── Picture helpers ──

function stringifyNvPicPr(
  id: number,
  name: string,
  opts: PictureOptions | undefined,
  ctx: WriteContext,
): string {
  // Click hyperlink (a:hlinkClick) lives inside p:cNvPr.
  let hlinkXml: string | undefined;
  if (opts?.hyperlink) {
    const key = registerHyperlink(opts.hyperlink, ctx);
    hlinkXml = buildHyperlinkElement("a:hlinkClick", opts.hyperlink, key);
  }
  const cNvPrXml = stringifyNonVisualDrawingProperties("p:cNvPr", id, opts, name, hlinkXml);
  const locks = opts?.locking ? (pictureLockingDesc.stringify(opts.locking, ctx) ?? "") : "";
  const preferRelativeResize =
    opts?.preferRelativeResize === undefined
      ? ""
      : ` preferRelativeResize="${opts.preferRelativeResize ? 1 : 0}"`;
  const cNvPicPr = locks
    ? `<p:cNvPicPr${preferRelativeResize}>${locks}</p:cNvPicPr>`
    : `<p:cNvPicPr${preferRelativeResize}/>`;
  return `<p:nvPicPr>${cNvPrXml}${cNvPicPr}${stringifyNvPr(opts ?? {})}</p:nvPicPr>`;
}

/** PPTX uses p:blipFill (not pic:blipFill). */
function stringifyAlternateBlipFillBranch(
  branch: AlternatePictureFillBranchOptions,
  ctx: WriteContext,
): string {
  const blipXml = stringify(
    blipDesc,
    {
      rawReferenceId: branch.relationshipId,
      compression: branch.compression,
      blipEffects: branch.blipEffects,
      ext: branch.blipExt,
    },
    ctx,
  );
  const sourceRectangle =
    branch.sourceRectangle === undefined
      ? ""
      : branch.sourceRectangle === null
        ? "<a:srcRect/>"
        : createSourceRectangle(branch.sourceRectangle);
  const fillRectangle =
    branch.fillRectangle === false
      ? "<a:stretch/>"
      : `<a:stretch>${createSourceRectangle(branch.fillRectangle ?? {}, "a:fillRect")}</a:stretch>`;
  return `<p:blipFill>${blipXml}${sourceRectangle}${fillRectangle}</p:blipFill>`;
}

function stringifyAlternateBlipFill(
  alternate: AlternatePictureBlipFillOptions,
  ctx: WriteContext,
): string {
  const prefix = alternate.choiceRequires;
  const choice = stringifyAlternateBlipFillBranch(alternate.choice, ctx);
  const fallback = stringifyAlternateBlipFillBranch(alternate.fallback, ctx);
  return (
    `<mc:AlternateContent><mc:Choice xmlns:${prefix}="${alternate.choiceRequiresNamespace}" ` +
    `Requires="${prefix}">${choice}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback>` +
    "</mc:AlternateContent>"
  );
}

function stringifyPptxBlipFill(
  blip: {
    /** Embedded media file name; undefined on a linked-only picture. */
    fileName?: string;
    /** External image source URL (a:blip @r:link). */
    sourceUrl?: string;
    /** Compression state (a:blip @cstate); absent = attribute omitted. */
    compression?: BlipCompression;
    /** Raw r:embed relationship ID for a reference without media bytes. */
    rawReferenceId?: string;
  },
  sourceRectangle?: SourceRectangleOptions,
  blipEffects?: PictureOptions["blipEffects"],
  ctx?: WriteContext,
  useLocalDpi?: boolean | "default",
  ext?: string,
  fillRectangle?: SourceRectangleOptions | false,
  rotWithShape?: boolean,
): string {
  // The linked source registers an {img-link:key} placeholder the compiler
  // rewrites per slide/layout while adding the External image relationship.
  // The placeholder body carries the "img-link:" prefix (hlink convention).
  const linkKey =
    blip.sourceUrl !== undefined && ctx
      ? `img-link:${(ctx as PptxWriteContext).addImageLink(blip.sourceUrl)}`
      : undefined;
  const blipXml = stringify(
    blipDesc,
    {
      referenceId: blip.fileName,
      rawReferenceId: blip.rawReferenceId,
      linkReferenceId: linkKey,
      compression: blip.compression,
      blipEffects,
      useLocalDpi,
      ext,
    },
    ctx,
  );
  const srcRect = sourceRectangle ? createSourceRectangle(sourceRectangle) : "";
  const fillRect =
    fillRectangle === false
      ? "<a:stretch/>"
      : `<a:stretch>${createSourceRectangle(fillRectangle ?? {}, "a:fillRect")}</a:stretch>`;
  const rootAttrs = rotWithShape ? ' rotWithShape="1"' : "";
  return `<p:blipFill${rootAttrs}>${blipXml}${srcRect}${fillRect}</p:blipFill>`;
}

function stringifyPicSpPr(opts: PictureOptions, ctx: WriteContext): string {
  const spPrContent = shapePropertiesDesc.stringify(
    {
      x: opts.x,
      y: opts.y,
      width: opts.width,
      height: opts.height,
      flipHorizontal: opts.flipHorizontal,
      flipVertical: opts.flipVertical,
      rotation: opts.rotation,
      // Pictures default to a rect frame; a parsed non-rect preset and an
      // explicit null (source omitted geometry) round-trip as written.
      geometry:
        opts.geometry === null || opts.customGeometry !== undefined
          ? undefined
          : (opts.geometry ?? "rect"),
      customGeometry: opts.customGeometry,
      fill: opts.fill,
      outline: opts.outline,
      effects: opts.effects,
      effectDag: opts.effectDag,
      scene3d: opts.scene3d,
      shape3d: opts.shape3d,
      extensions: opts.shapePropertiesExtensions,
      ext: opts.shapePropertiesExt,
    },
    ctx,
  );
  const bwAttr = opts.blackWhiteMode ? ` bwMode="${opts.blackWhiteMode}"` : "";
  if (!spPrContent) return `<p:spPr${bwAttr}/>`;
  return `<p:spPr${bwAttr}>${spPrContent}</p:spPr>`;
}

// ── Read helpers ──

/**
 * Read the non-visual drawing properties (id/name/description/title/hidden…)
 * from the first parent tag that holds a p:cNvPr (a:cNvPr tolerated); with no
 * parent tags, el itself is the non-visual properties container. Shared
 * probe+parse by every slide-child descriptor.
 */
export function readCnvPr(
  el: XmlElement,
  ...parentTags: string[]
): NonVisualDrawingPropertiesOptions & { id?: number } {
  const parents = parentTags.length ? parentTags.map((tag) => findChild(el, tag)) : [el];
  for (const parent of parents) {
    if (!parent) continue;
    const cNvPr = findChild(parent, "p:cNvPr") ?? findChild(parent, "a:cNvPr");
    if (!cNvPr) continue;
    const result: NonVisualDrawingPropertiesOptions & { id?: number } = {
      ...parseNonVisualDrawingProperties(cNvPr),
    };
    const id = attrNum(cNvPr, "id");
    if (id !== undefined) result.id = id;
    return result;
  }
  return {};
}

export function readNvSpPr(nvSpPr: XmlElement, ctx: ReadContext): ShapeOptions {
  const result: ShapeOptions = {};

  Object.assign(result, readCnvPr(nvSpPr));

  // Click hyperlink inside p:cNvPr (a:hlinkClick)
  const cNvPr = findChild(nvSpPr, "p:cNvPr") ?? findChild(nvSpPr, "a:cNvPr");
  const hlinkClick = cNvPr ? findChild(cNvPr, "a:hlinkClick") : undefined;
  if (hlinkClick) {
    result.hyperlink = readHyperlink(hlinkClick, ctx);
  }

  const nvPr = findChild(nvSpPr, "p:nvPr");
  if (nvPr) {
    if (nvPr.attributes) {
      if (nvPr.attributes["isPhoto"] !== undefined)
        result.isPhoto = parseOnOff(nvPr.attributes["isPhoto"]) ?? false;
      if (nvPr.attributes["userDrawn"] !== undefined)
        result.userDrawn = parseOnOff(nvPr.attributes["userDrawn"]) ?? false;
    }
    const ph = findChild(nvPr, "p:ph");
    if (ph?.attributes) {
      if (ph.attributes["type"] !== undefined)
        result.placeholder = xsdPlaceholderType.from(
          String(ph.attributes["type"]),
        ) as ShapeOptions["placeholder"];
      if (ph.attributes["idx"] !== undefined)
        result.placeholderIndex = Number(ph.attributes["idx"]);
      if (ph.attributes["sz"] !== undefined)
        result.placeholderSize = ph.attributes["sz"] as ShapeOptions["placeholderSize"];
      if (ph.attributes["orient"] !== undefined)
        result.placeholderOrientation = ph.attributes[
          "orient"
        ] as ShapeOptions["placeholderOrientation"];
      if (ph.attributes["hasCustomPrompt"] !== undefined)
        result.hasCustomPrompt = parseOnOff(ph.attributes["hasCustomPrompt"]) ?? false;
    } else if (ph) {
      result.placeholder = true;
    }
    result.customerData = parseCustDataLst(findChild(nvPr, "p:custDataLst"));
    const extLst = findChild(nvPr, "p:extLst");
    if (extLst) {
      result.nvPrExtensions = (extLst.elements ?? [])
        .filter((ext) => ext.name === "p:ext")
        .map((ext) => ({
          uri: attr(ext, "uri") ?? "",
          ...(ext.elements?.length
            ? { content: ext.elements.map((child) => stringifyElement(child)).join("") }
            : {}),
        }));
    }
  }

  const cNvSpPr = findChild(nvSpPr, "p:cNvSpPr");
  if (cNvSpPr) {
    const spLocks = findChild(cNvSpPr, "a:spLocks");
    if (spLocks) {
      result.locking = shapeLockingDesc.parse(spLocks, {} as ReadContext) as ShapeLockingOptions;
    }
    if (cNvSpPr.attributes?.["txBox"] !== undefined)
      result.textBox = parseOnOff(cNvSpPr.attributes["txBox"]) ?? false;
  }

  return result;
}

/** Parse p:spPr via the shared core descriptor: transform top-level, paint into
 * `properties` (fill absence marked null so stringify skips the fresh default). */
function readSpPr(
  spPr: XmlElement,
  ctx: ReadContext,
): Transform2DOptions & {
  properties: NonNullable<ShapeOptions["properties"]>;
} {
  const parsed = parse(shapePropertiesDesc, spPr, ctx);
  const { x, y, width, height, flipHorizontal, flipVertical, rotation, ...paint } = parsed;

  // A source spPr with no fill child inherits its fill — mark the absence so
  // stringify does not apply the fresh-authoring noFill default.
  if (paint.fill === undefined) paint.fill = null;
  if (paint.geometry === undefined) paint.geometry = null;

  // Collapse a preset geometry without adjustment values to the bare string
  // shorthand the public options accept.
  const geometry = paint.geometry;
  if (
    geometry !== null &&
    typeof geometry === "object" &&
    (!geometry.adjustmentValues || geometry.adjustmentValues.length === 0) &&
    geometry.preset
  ) {
    paint.geometry = geometry.preset;
  }

  return {
    x,
    y,
    width,
    height,
    flipHorizontal,
    flipVertical,
    rotation,
    properties: paint,
  };
}

/** Read x/y/width/height (in EMU) from an a:xfrm element. */
function readAlternateBlipFillBranch(
  fillEl: XmlElement,
  ctx: ReadContext,
): AlternatePictureFillBranchOptions {
  const blip = findChild(fillEl, "a:blip");
  const parsedBlip = blip ? parse(blipDesc, blip, ctx) : undefined;
  const sourceRectangleEl = findChild(fillEl, "a:srcRect");
  const stretchEl = findChild(fillEl, "a:stretch");
  const fillRectangleEl = stretchEl ? findChild(stretchEl, "a:fillRect") : undefined;
  return {
    relationshipId: blip ? (attr(blip, "r:embed") ?? "") : "",
    ...(parsedBlip?.compression !== undefined
      ? { compression: parsedBlip.compression as BlipCompression }
      : {}),
    ...(sourceRectangleEl
      ? { sourceRectangle: parse(sourceRectangleDesc, sourceRectangleEl, ctx) }
      : {}),
    ...(stretchEl
      ? {
          fillRectangle: fillRectangleEl ? parse(stretchDesc, stretchEl, ctx) : (false as const),
        }
      : {}),
    ...(parsedBlip?.blipEffects ? { blipEffects: parsedBlip.blipEffects } : {}),
    ...(parsedBlip?.ext !== undefined ? { blipExt: parsedBlip.ext } : {}),
  };
}

function readAlternateBlipFill(
  alternateEl: XmlElement,
  ctx: ReadContext,
): AlternatePictureBlipFillOptions | undefined {
  const choice = findChild(alternateEl, "mc:Choice");
  const fallback = findChild(alternateEl, "mc:Fallback");
  const choiceFill = choice ? findChild(choice, "p:blipFill") : undefined;
  const fallbackFill = fallback ? findChild(fallback, "p:blipFill") : undefined;
  if (!choice || !choiceFill || !fallbackFill) return undefined;
  const choiceRequires = attr(choice, "Requires");
  if (!choiceRequires) return undefined;
  return {
    choiceRequires,
    choiceRequiresNamespace: attr(choice, `xmlns:${choiceRequires}`) ?? "",
    choice: readAlternateBlipFillBranch(choiceFill, ctx),
    fallback: readAlternateBlipFillBranch(fallbackFill, ctx),
  };
}

export function readPositionFromXfrm(xfrm: XmlElement): Record<string, number> {
  const result: Record<string, number> = {};
  const off = findChild(xfrm, "a:off");
  if (off) {
    const x = attrNum(off, "x");
    if (x !== undefined) result.x = x;
    const y = attrNum(off, "y");
    if (y !== undefined) result.y = y;
  }
  const ext = findChild(xfrm, "a:ext");
  if (ext) {
    const cx = attrNum(ext, "cx");
    if (cx !== undefined) result.width = cx;
    const cy = attrNum(ext, "cy");
    if (cy !== undefined) result.height = cy;
  }
  return result;
}

// ── Image type helper ──
