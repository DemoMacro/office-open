/**
 * Presentation Properties (p:presentationPr) descriptor for PPTX.
 *
 * @module
 */

import { parseOnOff } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { findChild } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";
import {
  buildPresentationPropertiesXml,
  type PresentationPropertiesOptions,
} from "@parts/presentation-properties";
import type { ShowOptions } from "@shared/file";

// ── Descriptor ──

export const presentationPropertiesDesc: CustomDescriptor<PresentationPropertiesOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    return buildPresentationPropertiesXml(opts);
  },

  parse(el, _ctx) {
    return parsePresentationProperties(el);
  },
};

// ── Parse ──

function parsePresentationProperties(el: XmlElement): PresentationPropertiesOptions {
  const result: Partial<PresentationPropertiesOptions> = {};

  // show (p:showPr in real PPTX files, or p:show for round-trip compat)
  const showPr = findChild(el, "p:showPr") ?? findChild(el, "p:show");
  if (showPr) {
    const showOpts: Partial<ShowOptions> = {};
    if (showPr.attributes?.["loop"] !== undefined)
      showOpts.loop = parseOnOff(showPr.attributes["loop"]) ?? false;
    if (showPr.attributes?.["showNarration"] !== undefined)
      showOpts.showNarration = parseOnOff(showPr.attributes["showNarration"]) ?? false;
    if (showPr.attributes?.["useTimings"] !== undefined)
      showOpts.useTimings = parseOnOff(showPr.attributes["useTimings"]) ?? false;
    // Show type from child elements (kiosk/browse/present)
    if (findChild(showPr, "p:kiosk")) {
      showOpts.type = "kiosk";
    } else if (findChild(showPr, "p:browse")) {
      showOpts.type = "browse";
    } else if (findChild(showPr, "p:present")) {
      showOpts.type = "present";
    }
    if (Object.keys(showOpts).length > 0) result.show = showOpts as ShowOptions;
  }

  return result as PresentationPropertiesOptions;
}
