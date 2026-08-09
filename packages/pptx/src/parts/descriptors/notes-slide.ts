/**
 * Notes Slide (p:notes) descriptor for PPTX.
 *
 * CT_NotesSlide is CT_Slide minus transition/timing: cSld (bg + spTree) +
 * clrMapOvr (EG_ChildSlide) + extLst. The shape tree holds the slide-image
 * placeholder (p:ph type="sldImg", no txBody) and the notes body placeholder,
 * plus any extra shapes a user adds. Structured stringify mirrors slideDesc so
 * a parsed notes slide round-trips byte-exact; the simple `text` shorthand
 * keeps the fresh API (text-only notes) one line.
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { findChild } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import { buildNotesSlideXml } from "@parts/notes-slide";
import type { SlideChild } from "@parts/slide/slide-child";
import { SP_TREE_HEADER } from "@shared/constants";

import type { PptxWriteContext } from "../../context";
import { backgroundDesc, type BackgroundDescriptorOptions } from "./background";
import { stringifyChild, parseChild } from "./bridge";
import { colorMapOverrideDesc, type ColorMapOverrideOptions } from "./color-map-override";

// ── Types ──

export interface NotesSlideDescriptorOptions {
  /** Notes text — shorthand that builds a fresh sldImg + body placeholder slide. */
  text?: string;
  /** Shape-tree children (structured). When set, the structured path is used. */
  children?: SlideChild[];
  background?: BackgroundDescriptorOptions;
  colorMapOverride?: ColorMapOverrideOptions;
  showMasterSp?: boolean;
  showMasterPhAnim?: boolean;
  /** Hidden notes slide (p:notes/@show="0"). */
  show?: boolean;
}

// ── Descriptor ──

export const notesSlideDesc: CustomDescriptor<NotesSlideDescriptorOptions, PptxWriteContext> = {
  kind: "custom",

  stringify(opts, ctx) {
    // Simple text path — fresh builder with sldImg + body placeholders.
    const structured =
      opts.children !== undefined ||
      opts.background !== undefined ||
      opts.colorMapOverride !== undefined;
    if (!structured) return buildNotesSlideXml({ text: opts.text });

    // Structured path — mirror slideDesc (p:notes has no transition/timing).
    const parts: string[] = [];
    const attrs: string[] = [];
    if (opts.showMasterSp === false) attrs.push(' showMasterSp="0"');
    if (opts.showMasterPhAnim === false) attrs.push(' showMasterPhAnim="0"');
    if (opts.show === false) attrs.push(' show="0"');
    parts.push(
      '<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"' +
        attrs.join("") +
        ">",
    );

    parts.push("<p:cSld>");
    if (opts.background) parts.push(backgroundDesc.stringify(opts.background, ctx) ?? "");
    parts.push("<p:spTree>");
    parts.push(SP_TREE_HEADER);
    if (opts.children) {
      for (const child of opts.children) {
        const xml = stringifyChild(child, ctx);
        if (xml) parts.push(xml);
      }
    }
    parts.push("</p:spTree>");
    parts.push("</p:cSld>");

    parts.push(
      colorMapOverrideDesc.stringify(opts.colorMapOverride, ctx) ??
        "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>",
    );

    parts.push("</p:notes>");
    return parts.join("");
  },

  parse(el, ctx) {
    const result: Partial<NotesSlideDescriptorOptions> = {};

    if (el.attributes) {
      if (el.attributes["showMasterSp"] !== undefined)
        result.showMasterSp = String(el.attributes["showMasterSp"]) !== "0";
      if (el.attributes["showMasterPhAnim"] !== undefined)
        result.showMasterPhAnim = String(el.attributes["showMasterPhAnim"]) !== "0";
      if (el.attributes["show"] !== undefined) result.show = String(el.attributes["show"]) !== "0";
    }

    const cSld = findChild(el, "p:cSld");
    if (cSld) {
      const bg = findChild(cSld, "p:bg");
      if (bg) result.background = backgroundDesc.parse(bg, ctx);

      const spTree = findChild(cSld, "p:spTree");
      if (spTree) {
        const children: SlideChild[] = [];
        for (const child of spTree.elements ?? []) {
          // Skip the tree container structure (nvGrpSpPr/grpSpPr).
          if (child.name === "p:nvGrpSpPr" || child.name === "p:grpSpPr") continue;
          const parsed = parseChild(child as Element, ctx);
          if (parsed !== undefined) children.push(parsed as SlideChild);
        }
        if (children.length > 0) result.children = children;
      }
    }

    const clrMapOvr = findChild(el, "p:clrMapOvr");
    if (clrMapOvr) result.colorMapOverride = colorMapOverrideDesc.parse(clrMapOvr, ctx);

    return result as NotesSlideDescriptorOptions;
  },
};
