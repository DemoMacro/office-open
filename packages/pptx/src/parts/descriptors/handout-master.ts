/**
 * Handout Master (p:handoutMaster) descriptor for PPTX.
 *
 * @module
 */

import { parseColorMapping, stringifyColorMapping } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { findChild, stringify } from "@office-open/xml";
import { buildHfAttrs, buildHandoutMasterXml, parseHeaderFooter } from "@parts/handout-master";
import type { HandoutMasterOptions } from "@parts/handout-master";
import { NS } from "@parts/slide-layout";
import { buildBackgroundXml } from "@parts/slide-master";
import type { SlideChild } from "@parts/slide/slide-child";
import { SP_TREE_HEADER } from "@shared/constants";

import type { PptxWriteContext } from "../../context";
import { backgroundDesc } from "./background";
import { parseChild, stringifyChild } from "./bridge";
import { withChildId } from "./slide-master";

// ── Types ──

export interface HandoutMasterDescriptorOptions {
  options?: HandoutMasterOptions;
}

// ── Descriptor ──

export const handoutMasterDesc: CustomDescriptor<HandoutMasterDescriptorOptions, PptxWriteContext> =
  {
    kind: "custom",

    stringify(opts, ctx) {
      // A parsed source carries its spTree shapes; a fresh one keeps the
      // minimal Office default.
      if (opts.options?.children?.length) {
        const parts: string[] = [];
        parts.push(`<p:handoutMaster ${NS}>`);
        parts.push("<p:cSld>");
        parts.push(buildBackgroundXml(opts.options.background));
        parts.push("<p:spTree>");
        parts.push(SP_TREE_HEADER);
        let childId = 100;
        for (const child of opts.options.children) {
          const xml = stringifyChild(withChildId(child, childId++), ctx);
          if (xml) parts.push(xml);
        }
        parts.push("</p:spTree>");
        if (opts.options.ext !== undefined) parts.push(`<p:extLst>${opts.options.ext}</p:extLst>`);
        parts.push("</p:cSld>");
        parts.push(stringifyColorMapping(opts.options.colorMapping, "p:clrMap"));
        if (opts.options.headerFooter !== undefined) {
          parts.push(`<p:hf ${buildHfAttrs(opts.options.headerFooter)}/>`);
        }
        parts.push("</p:handoutMaster>");
        return parts.join("");
      }
      return buildHandoutMasterXml(opts.options);
    },

    parse(el, ctx) {
      const options: HandoutMasterOptions = {};

      const cSld = findChild(el, "p:cSld");
      if (cSld) {
        const bg = findChild(cSld, "p:bg");
        if (bg) {
          const bgOpts = backgroundDesc.parse(bg, ctx);
          if (bgOpts && Object.keys(bgOpts).length > 0) options.background = bgOpts;
        }

        const spTree = findChild(cSld, "p:spTree");
        if (spTree) {
          const children: SlideChild[] = [];
          for (const child of spTree.elements ?? []) {
            if (child.name === "p:nvGrpSpPr" || child.name === "p:grpSpPr") continue;
            const parsed = parseChild(child, ctx);
            if (parsed !== undefined) children.push(parsed);
          }
          if (children.length > 0) options.children = children;
        }
      }

      const colorMapping = parseColorMapping(findChild(el, "p:clrMap"));
      if (colorMapping) options.colorMapping = colorMapping;

      const headerFooter = parseHeaderFooter(findChild(el, "p:hf"));
      if (headerFooter) options.headerFooter = headerFooter;

      const extLst = cSld && findChild(cSld, "p:extLst");
      if (extLst) options.ext = stringify(extLst);

      return Object.keys(options).length > 0 ? { options } : {};
    },
  };
