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

import { parseOnOff } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { parseNonVisualDrawingProperties } from "@office-open/core/drawing";
import { findChild, stringify as stringifyXml } from "@office-open/xml";
import { attr, attrNum } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import { buildNotesSlideXml } from "@parts/notes-slide";
import type { SlideChild } from "@parts/slide/slide-child";
import { stringifySpTreeHeader } from "@shared/constants";

import type { PptxWriteContext } from "../../context";
import type { BackgroundOptions } from "../background";
import { backgroundDesc } from "./background";
import { stringifyChild, parseChild } from "./bridge";
import { colorMappingOverrideDesc, type ColorMappingOverrideOptions } from "./color-map-override";

// ── Types ──

export interface NotesSlideOptions {
  /** Root spTree cNvPr id — round-trip only. */
  shapeTreeId?: number;
  /** Root spTree cNvPr name — round-trip only. */
  shapeTreeName?: string;
  /** Root spTree cNvPr creation id — round-trip only. */
  shapeTreeCreationId?: string;
  /** Root spTree cNvPr extension-list inner XML — round-trip only. */
  shapeTreeExt?: string;
  /** Root spTree group properties black-and-white mode (`@bwMode`). */
  shapeTreeBlackWhiteMode?: string;
  /** Omit the shape-tree root transform when the source omitted it. */
  groupTransformless?: boolean;
  /** Notes text — shorthand that builds a fresh sldImg + body placeholder slide. */
  text?: string;
  /** Shape-tree children (structured). When set, the structured path is used. */
  children?: SlideChild[];
  background?: BackgroundOptions;
  colorMappingOverride?: ColorMappingOverrideOptions;
  /**
   * Raw inner XML of the p:extLst inside p:cSld (CT_CommonSlideData tail —
   * where p14:creationId lives) — verbatim round-trip.
   */
  cSldExt?: string;
  showMasterShapes?: boolean;
  showMasterPlaceholderAnimations?: boolean;
  /** Hidden notes slide (p:notes/`@show`="0"). */
  show?: boolean;
}

// ── Descriptor ──

export const notesSlideDesc: CustomDescriptor<NotesSlideOptions, PptxWriteContext> = {
  kind: "custom",

  stringify(opts, ctx) {
    // Simple text path — fresh builder with sldImg + body placeholders.
    const structured =
      opts.children !== undefined ||
      opts.shapeTreeId !== undefined ||
      opts.shapeTreeName !== undefined ||
      opts.shapeTreeCreationId !== undefined ||
      opts.shapeTreeExt !== undefined ||
      opts.shapeTreeBlackWhiteMode !== undefined ||
      opts.groupTransformless !== undefined ||
      opts.background !== undefined ||
      opts.colorMappingOverride !== undefined ||
      opts.cSldExt !== undefined;
    if (!structured) return buildNotesSlideXml({ text: opts.text });

    // Structured path — mirror slideDesc (p:notes has no transition/timing).
    const parts: string[] = [];
    const attrs: string[] = [];
    if (opts.showMasterShapes === false) attrs.push(' showMasterSp="0"');
    if (opts.showMasterPlaceholderAnimations === false) attrs.push(' showMasterPhAnim="0"');
    if (opts.show === false) attrs.push(' show="0"');
    parts.push(
      '<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"' +
        attrs.join("") +
        ">",
    );

    parts.push("<p:cSld>");
    if (opts.background) parts.push(backgroundDesc.stringify(opts.background, ctx) ?? "");
    parts.push("<p:spTree>");
    parts.push(
      stringifySpTreeHeader(
        opts.shapeTreeId,
        opts.shapeTreeName,
        opts.shapeTreeBlackWhiteMode,
        opts.groupTransformless,
        opts.shapeTreeCreationId,
        opts.shapeTreeExt,
      ),
    );
    if (opts.children) {
      for (const child of opts.children) {
        const xml = stringifyChild(child, ctx);
        if (xml) parts.push(xml);
      }
    }
    parts.push("</p:spTree>");
    // cSld-tail extLst (p14:creationId's home) — verbatim.
    if (opts.cSldExt) parts.push(`<p:extLst>${opts.cSldExt}</p:extLst>`);
    parts.push("</p:cSld>");

    parts.push(colorMappingOverrideDesc.stringify(opts.colorMappingOverride, ctx) ?? "");

    parts.push("</p:notes>");
    return parts.join("");
  },

  parse(el, ctx) {
    const result: Partial<NotesSlideOptions> = {};

    if (el.attributes) {
      if (el.attributes["showMasterSp"] !== undefined)
        result.showMasterShapes = parseOnOff(String(el.attributes["showMasterSp"])) ?? true;
      if (el.attributes["showMasterPhAnim"] !== undefined)
        result.showMasterPlaceholderAnimations =
          parseOnOff(String(el.attributes["showMasterPhAnim"])) ?? true;
      if (el.attributes["show"] !== undefined)
        result.show = parseOnOff(String(el.attributes["show"])) ?? true;
    }

    const cSld = findChild(el, "p:cSld");
    if (cSld) {
      const bg = findChild(cSld, "p:bg");
      if (bg) result.background = backgroundDesc.parse(bg, ctx);

      const spTree = findChild(cSld, "p:spTree");
      if (spTree) {
        const shapeTreeGroupProperties = findChild(spTree, "p:grpSpPr");
        if (
          shapeTreeGroupProperties &&
          findChild(shapeTreeGroupProperties, "a:xfrm") === undefined
        ) {
          result.groupTransformless = true;
        }
        const shapeTreeBwMode = attr(shapeTreeGroupProperties, "bwMode");
        if (shapeTreeBwMode !== undefined) result.shapeTreeBlackWhiteMode = shapeTreeBwMode;
        const rootGroupProperties = findChild(spTree, "p:nvGrpSpPr");
        const rootDrawingProperties = findChild(rootGroupProperties, "p:cNvPr");
        const rootProperties = parseNonVisualDrawingProperties(rootDrawingProperties);
        const shapeTreeId = rootDrawingProperties
          ? attrNum(rootDrawingProperties, "id")
          : undefined;
        if (shapeTreeId !== undefined) result.shapeTreeId = shapeTreeId;
        if (rootProperties.name !== undefined) result.shapeTreeName = rootProperties.name;
        if (rootProperties.creationId !== undefined)
          result.shapeTreeCreationId = rootProperties.creationId;
        if (rootProperties.ext !== undefined) result.shapeTreeExt = rootProperties.ext;
        const children: SlideChild[] = [];
        for (const child of spTree.elements ?? []) {
          // Skip the tree container structure (nvGrpSpPr/grpSpPr).
          if (child.name === "p:nvGrpSpPr" || child.name === "p:grpSpPr") continue;
          const parsed = parseChild(child as Element, ctx);
          if (parsed !== undefined) children.push(parsed as SlideChild);
        }
        if (children.length > 0) result.children = children;
      }

      // cSld-tail extLst (p14:creationId's home) — verbatim.
      const cSldExtLst = findChild(cSld, "p:extLst");
      if (cSldExtLst) {
        const inner = stringifyXml(cSldExtLst);
        if (inner) result.cSldExt = inner;
      }
    }

    const clrMapOvr = findChild(el, "p:clrMapOvr");
    if (clrMapOvr) result.colorMappingOverride = colorMappingOverrideDesc.parse(clrMapOvr, ctx);

    return result as NotesSlideOptions;
  },
};
