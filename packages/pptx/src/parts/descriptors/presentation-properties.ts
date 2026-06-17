/**
 * Presentation Properties (p:presentationPr) descriptor for PPTX.
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { findChild } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";
import {
  buildPresPropsXml,
  type PresentationPropertiesFullOptions,
} from "@parts/presentation-properties";

// ── Types ──

export type PresPropsDescriptorOptions = PresentationPropertiesFullOptions;

// ── Descriptor ──

export const presPropsDesc: CustomDescriptor<PresPropsDescriptorOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    return buildPresPropsXml(opts);
  },

  parse(el, _ctx) {
    return parsePresProps(el);
  },
};

// ── Parse ──

function parsePresProps(el: XmlElement): PresPropsDescriptorOptions {
  const result: Record<string, unknown> = {};

  // show (p:showPr in real PPTX files, or p:show for round-trip compat)
  const showPr = findChild(el, "p:showPr") ?? findChild(el, "p:show");
  if (showPr) {
    const showOpts: Record<string, unknown> = {};
    if (showPr.attributes?.["loop"] !== undefined)
      showOpts.loop = showPr.attributes["loop"] === "1";
    if (showPr.attributes?.["showNarration"] !== undefined)
      showOpts.showNarration = showPr.attributes["showNarration"] === "1";
    if (showPr.attributes?.["useTimings"] !== undefined)
      showOpts.useTimings = showPr.attributes["useTimings"] === "1";
    // Show type from child elements (kiosk/browse/present)
    if (findChild(showPr, "p:kiosk")) {
      showOpts.type = "kiosk";
    } else if (findChild(showPr, "p:browse")) {
      showOpts.type = "browse";
    } else if (findChild(showPr, "p:present")) {
      showOpts.type = "present";
    }
    if (Object.keys(showOpts).length > 0) result.show = showOpts;
  }

  return result as unknown as PresPropsDescriptorOptions;
}
