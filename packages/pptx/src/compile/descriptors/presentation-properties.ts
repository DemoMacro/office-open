/**
 * Presentation Properties (p:presentationPr) descriptor for PPTX.
 *
 * @module
 */

import {
  buildPresPropsXml,
  type PresentationPropertiesFullOptions,
} from "@file/presentation-properties";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { findChild } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";

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

function parsePresProps(el: XmlElement): Partial<PresPropsDescriptorOptions> {
  const result: Record<string, unknown> = {};

  // show (p:show)
  const show = findChild(el, "p:show");
  if (show?.attributes) {
    const showOpts: Record<string, unknown> = {};
    if (show.attributes["loop"] !== undefined) showOpts.loop = show.attributes["loop"] === "1";
    if (show.attributes["startSlideNum"] !== undefined)
      showOpts.startSlideNum = Number(show.attributes["startSlideNum"]);
    if (show.attributes["showNarration"] !== undefined)
      showOpts.showNarration = show.attributes["showNarration"] === "1";
    if (show.attributes["animation"] !== undefined)
      showOpts.animation = show.attributes["animation"] === "1";
    if (show.attributes["advanceMode"]) showOpts.advanceMode = show.attributes["advanceMode"];
    if (show.attributes["present"]) showOpts.present = show.attributes["present"] === "1";
    if (Object.keys(showOpts).length > 0) result.show = showOpts;
  }

  return result as Partial<PresPropsDescriptorOptions>;
}
