/**
 * Slide (p:sld) descriptor for PPTX.
 *
 * @module
 */

import { SP_TREE_HEADER } from "@file/constants";
import type { SlideChild as LegacySlideChild } from "@file/slide/slide-child";
import type { CustomDescriptor, WriteContext } from "@office-open/core/descriptor";
import { attr, attrNum, findChild } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";

import { parseTiming } from "../../parse/animation";
import { parseChild } from "./bridge";
import { shapeDesc, pictureDesc } from "./shape";
import type {
  ShapeDescriptorOptions,
  PictureDescriptorOptions,
  TextBodyDescriptorOptions,
} from "./shape";

// ── Types ──

export interface HeaderFooterDescriptorOptions {
  slideNumber?: boolean;
  dateTime?: boolean;
  footer?: boolean;
  header?: boolean;
}

export interface AnimationDescriptorOptions {
  shapeId: number;
  options: unknown;
}

export interface SlideDescriptorOptions {
  children?: readonly SlideChild[];
  background?: BackgroundDescriptorOptions;
  transition?: TransitionDescriptorOptions;
  showMasterSp?: boolean;
  showMasterPhAnim?: boolean;
  controls?: readonly ControlDescriptorOptions[];
  customerData?: readonly { readonly rId: string }[];
  headerFooter?: HeaderFooterDescriptorOptions;
  animations?: readonly AnimationDescriptorOptions[];
}

/** Discriminated union for slide children (JSON-friendly). */
export type SlideChild =
  | { shape: ShapeDescriptorOptions }
  | { picture: PictureDescriptorOptions }
  | { text: TextBodyDescriptorOptions }
  | { contentPart: { readonly rId: string } };

export interface BackgroundDescriptorOptions {
  color?: string;
  transparency?: number;
}

export interface TransitionDescriptorOptions {
  type?:
    | "none"
    | "fade"
    | "push"
    | "wipe"
    | "split"
    | "cover"
    | "pull"
    | "dissolve"
    | "wheel"
    | "random";
  speed?: "slow" | "medium" | "fast";
  advanceOnClick?: boolean;
  advanceAfterMs?: number;
}

export interface ControlDescriptorOptions {
  shapeId?: number;
  name?: string;
  showAsIcon?: boolean;
  rId?: string;
  imageWidth?: number;
  imageHeight?: number;
}

// ── Slide (p:sld) descriptor ──

export const slideDesc: CustomDescriptor<SlideDescriptorOptions> = {
  kind: "custom",

  stringify(opts, ctx) {
    const parts: string[] = [];

    // Opening tag with namespace declarations
    const sldAttrs: string[] = [];
    if (opts.showMasterSp === false) sldAttrs.push(' showMasterSp="0"');
    if (opts.showMasterPhAnim === false) sldAttrs.push(' showMasterPhAnim="0"');
    parts.push(
      `<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"${sldAttrs.join("")}>`,
    );

    // p:cSld — common slide data
    parts.push("<p:cSld>");

    if (opts.background) {
      parts.push(stringifyBackground(opts.background));
    }

    // p:spTree — shape tree
    parts.push("<p:spTree>");
    parts.push(SP_TREE_HEADER);

    if (opts.children) {
      for (const child of opts.children) {
        const xml = stringifySlideChild(child, ctx);
        if (xml) parts.push(xml);
      }
    }

    parts.push("</p:spTree>");

    // custDataLst
    if (opts.customerData && opts.customerData.length > 0) {
      const cdItems = opts.customerData.map((d) => `<p:custData r:id="${d.rId}"/>`).join("");
      parts.push(`<p:custDataLst>${cdItems}</p:custDataLst>`);
    }

    // controls
    if (opts.controls && opts.controls.length > 0) {
      const ctrlItems = opts.controls
        .map((c) => {
          const attrs: string[] = [];
          if (c.shapeId !== undefined) attrs.push(`spid="${c.shapeId}"`);
          if (c.name) attrs.push(`name="${c.name}"`);
          if (c.showAsIcon) attrs.push('showAsIcon="1"');
          if (c.rId) attrs.push(`r:id="${c.rId}"`);
          if (c.imageWidth !== undefined) attrs.push(`imgW="${c.imageWidth}"`);
          if (c.imageHeight !== undefined) attrs.push(`imgH="${c.imageHeight}"`);
          return `<p:control ${attrs.join(" ")}/>`;
        })
        .join("");
      parts.push(`<p:controls>${ctrlItems}</p:controls>`);
    }

    parts.push("</p:cSld>");

    // p:clrMapOvr
    parts.push("<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>");

    // p:transition (optional)
    if (opts.transition) {
      parts.push(stringifyTransition(opts.transition));
    }

    parts.push("</p:sld>");
    return parts.join("");
  },

  parse(el, _ctx) {
    const result: Record<string, unknown> = {};

    // Root attributes
    if (el.attributes) {
      if (el.attributes["showMasterSp"] !== undefined)
        result.showMasterSp = el.attributes["showMasterSp"] !== "0";
      if (el.attributes["showMasterPhAnim"] !== undefined)
        result.showMasterPhAnim = el.attributes["showMasterPhAnim"] !== "0";
    }

    // p:cSld
    const cSld = findChild(el, "p:cSld");
    if (cSld) {
      // Background
      const bg = findChild(cSld, "p:bg");
      if (bg) result.background = readBackground(bg);

      // Shape tree
      const spTree = findChild(cSld, "p:spTree");
      if (spTree) {
        const children: LegacySlideChild[] = [];
        if (spTree.elements) {
          for (const child of spTree.elements) {
            // Skip tree container structure
            if (child.name === "p:nvGrpSpPr" || child.name === "p:grpSpPr") continue;
            const parsed = parseChild(child, _ctx);
            if (parsed !== undefined) children.push(parsed);
          }
        }
        if (children.length > 0) result.children = children;
      }

      // Header/Footer (p:hf)
      const hf = findChild(cSld, "p:hf");
      if (hf) {
        const hfOpts: HeaderFooterDescriptorOptions = {};
        if (findChild(hf, "p:sldNum")) hfOpts.slideNumber = true;
        if (findChild(hf, "p:dt")) hfOpts.dateTime = true;
        if (findChild(hf, "p:ftr")) hfOpts.footer = true;
        if (findChild(hf, "p:hdr")) hfOpts.header = true;
        if (hfOpts.slideNumber || hfOpts.dateTime || hfOpts.footer || hfOpts.header)
          result.headerFooter = hfOpts;
      }
    }

    // p:transition
    const transition = findChild(el, "p:transition");
    if (transition) result.transition = readTransition(transition);

    // p:timing → animations
    const timing = findChild(el, "p:timing");
    if (timing) {
      const animMap = parseTiming(timing);
      if (animMap.size > 0) {
        const animations: AnimationDescriptorOptions[] = [];
        for (const [shapeId, options] of animMap) {
          animations.push({ shapeId, options });
        }
        result.animations = animations;
      }
    }

    // custDataLst
    const custDataLst = findChild(el, "p:custDataLst");
    if (custDataLst) {
      const items: { rId: string }[] = [];
      for (const cd of custDataLst.elements ?? []) {
        if (cd.name === "p:custData") {
          const rId = attr(cd, "r:id");
          if (rId) items.push({ rId });
        }
      }
      if (items.length > 0) result.customerData = items;
    }

    // controls
    const controls = findChild(el, "p:controls");
    if (controls) {
      const items: ControlDescriptorOptions[] = [];
      for (const ctrl of controls.elements ?? []) {
        if (ctrl.name !== "p:control") continue;
        const item: ControlDescriptorOptions = {};
        const spid = attrNum(ctrl, "spid");
        if (spid !== undefined) item.shapeId = spid;
        const name = attr(ctrl, "name");
        if (name) item.name = name;
        if (attr(ctrl, "showAsIcon") === "1") item.showAsIcon = true;
        const rId = attr(ctrl, "r:id");
        if (rId) item.rId = rId;
        const imgW = attrNum(ctrl, "imgW");
        if (imgW !== undefined) item.imageWidth = imgW;
        const imgH = attrNum(ctrl, "imgH");
        if (imgH !== undefined) item.imageHeight = imgH;
        items.push(item);
      }
      if (items.length > 0) result.controls = items;
    }

    return result as Partial<SlideDescriptorOptions>;
  },
};

// ── Child serializer ──

function stringifySlideChild(child: SlideChild, ctx: WriteContext): string | undefined {
  if ("shape" in child) return shapeDesc.stringify(child.shape, ctx);
  if ("picture" in child) return pictureDesc.stringify(child.picture, ctx);
  if ("contentPart" in child) return `<p:contentPart r:id="${child.contentPart.rId}"/>`;
  return undefined;
}

// ── Background helpers ──

function stringifyBackground(opts: BackgroundDescriptorOptions): string {
  if (opts.color) {
    const alphaAttr =
      opts.transparency !== undefined
        ? `><a:srgbClr val="${opts.color.replace("#", "")}"><a:alpha val="${Math.round((100 - opts.transparency) * 1000)}"/></a:srgbClr></p:bgPr`
        : `><a:solidFill><a:srgbClr val="${opts.color.replace("#", "")}"/></a:solidFill></p:bgPr`;
    return `<p:bg><p:bgPr${alphaAttr}></p:bg>`;
  }
  return "<p:bg/>";
}

function readBackground(bg: XmlElement): BackgroundDescriptorOptions {
  const result: BackgroundDescriptorOptions = {};
  const bgPr = findChild(bg, "p:bgPr");
  if (bgPr) {
    const solidFill = findChild(bgPr, "a:solidFill");
    if (solidFill) {
      const srgbClr = findChild(solidFill, "a:srgbClr");
      if (srgbClr?.attributes?.["val"]) {
        result.color = String(srgbClr.attributes["val"]);
        const alpha = findChild(srgbClr, "a:alpha");
        if (alpha?.attributes?.["val"]) {
          result.transparency = 100 - Number(alpha.attributes["val"]) / 1000;
        }
      }
    }
  }
  return result;
}

// ── Transition helpers ──

function stringifyTransition(opts: TransitionDescriptorOptions): string {
  const parts: string[] = [];

  if (opts.type === "none") return "";
  if (opts.type === "fade") parts.push("<p:fade/>");
  else if (opts.type === "push") parts.push('<p:push dir="l"/>');
  else if (opts.type === "wipe") parts.push('<p:wipe dir="d"/>');
  else if (opts.type === "split") parts.push('<p:split orient="horz"/>');
  else if (opts.type === "cover") parts.push('<p:cover dir="l"/>');
  else if (opts.type === "pull") parts.push('<p:pull dir="l"/>');
  else if (opts.type === "dissolve") parts.push("<p:dissolve/>");
  else if (opts.type === "wheel") parts.push('<p:wheel spokes="4"/>');
  else if (opts.type === "random") parts.push("<p:random/>");

  const attrs: string[] = [];
  if (opts.speed) attrs.push(`spd="${opts.speed}"`);
  if (opts.advanceOnClick !== undefined) attrs.push(`advClick="${opts.advanceOnClick ? 1 : 0}"`);
  if (opts.advanceAfterMs !== undefined) attrs.push(`advTm="${opts.advanceAfterMs}"`);

  const attrStr = attrs.length ? " " + attrs.join(" ") : "";
  const body = parts.join("");
  return body ? `<p:transition${attrStr}>${body}</p:transition>` : `<p:transition${attrStr}/>`;
}

function readTransition(el: XmlElement): TransitionDescriptorOptions {
  const result: TransitionDescriptorOptions = {};

  if (el.attributes) {
    if (el.attributes["spd"] !== undefined)
      result.speed = el.attributes["spd"] as "slow" | "medium" | "fast";
    if (el.attributes["advClick"] !== undefined)
      result.advanceOnClick = el.attributes["advClick"] === "1";
    if (el.attributes["advTm"] !== undefined)
      result.advanceAfterMs = Number(el.attributes["advTm"]);
  }

  // Detect transition type from child elements
  if (findChild(el, "p:fade")) result.type = "fade";
  else if (findChild(el, "p:push")) result.type = "push";
  else if (findChild(el, "p:wipe")) result.type = "wipe";
  else if (findChild(el, "p:split")) result.type = "split";
  else if (findChild(el, "p:cover")) result.type = "cover";
  else if (findChild(el, "p:pull")) result.type = "pull";
  else if (findChild(el, "p:dissolve")) result.type = "dissolve";
  else if (findChild(el, "p:wheel")) result.type = "wheel";
  else if (findChild(el, "p:random")) result.type = "random";

  return result;
}
