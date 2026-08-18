/**
 * Presentation Properties (p:presentationPr) descriptor for PPTX.
 *
 * @module
 */

import { parseOnOff } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, attrNum, findChild, stringify as stringifyXml } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";
import {
  buildPresentationPropertiesXml,
  type PresentationPropertiesOptions,
  type PrintPropertiesOptions,
  type WebPropertiesOptions,
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

const COLOR_MODE_FROM_XSD: Record<string, PrintPropertiesOptions["colorMode"]> = {
  bw: "blackWhite",
  gray: "gray",
  clr: "color",
};

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
    if (showPr.attributes?.["showAnimation"] !== undefined)
      showOpts.showAnimation = parseOnOff(showPr.attributes["showAnimation"]) ?? false;
    if (showPr.attributes?.["useTimings"] !== undefined)
      showOpts.useTimings = parseOnOff(showPr.attributes["useTimings"]) ?? false;
    // Show type from child elements (kiosk/browse/present)
    if (findChild(showPr, "p:kiosk")) {
      showOpts.type = "kiosk";
      const restart = attrNum(findChild(showPr, "p:kiosk")!, "restart");
      if (restart !== undefined) showOpts.restart = restart;
    } else if (findChild(showPr, "p:browse")) {
      showOpts.type = "browse";
      if (String(attr(findChild(showPr, "p:browse")!, "showScrollbar")) === "0") {
        showOpts.showScrollbar = false;
      }
    } else if (findChild(showPr, "p:present")) {
      showOpts.type = "present";
    }
    const sldRg = findChild(showPr, "p:sldRg");
    if (sldRg) {
      const start = attrNum(sldRg, "st");
      const end = attrNum(sldRg, "end");
      if (start !== undefined && end !== undefined) showOpts.slideRange = { start, end };
    }
    const penClr = findChild(showPr, "p:penClr");
    if (penClr) {
      const srgb = findChild(penClr, "a:srgbClr");
      const val = srgb ? attr(srgb, "val") : undefined;
      if (val) showOpts.penColor = val;
    }
    if (Object.keys(showOpts).length > 0) result.show = showOpts as ShowOptions;
  }

  // web (p:webPr) — default-true booleans only carry "0"
  const webPr = findChild(el, "p:webPr");
  if (webPr) {
    const web: Partial<WebPropertiesOptions> = {};
    if (parseOnOff(webPr.attributes?.["showAnimation"])) web.showAnimation = true;
    if (String(webPr.attributes?.["resizeGraphics"]) === "0") web.resizeGraphics = false;
    if (parseOnOff(webPr.attributes?.["allowPng"])) web.allowPng = true;
    if (parseOnOff(webPr.attributes?.["relyOnVml"])) web.relyOnVml = true;
    if (String(webPr.attributes?.["organizeInFolders"]) === "0") web.organizeInFolders = false;
    if (String(webPr.attributes?.["useLongFilenames"]) === "0") web.useLongFilenames = false;
    const imageSize = attr(webPr, "imgSz");
    if (imageSize) web.imageSize = imageSize;
    const encoding = attr(webPr, "encoding");
    if (encoding) web.encoding = encoding;
    const color = attr(webPr, "clr");
    if (color) web.color = color;
    if (Object.keys(web).length > 0) result.web = web as WebPropertiesOptions;
  }

  // print (p:prnPr)
  const prnPr = findChild(el, "p:prnPr");
  if (prnPr) {
    const print: Partial<PrintPropertiesOptions> = {};
    const printWhat = attr(prnPr, "prnWhat");
    if (printWhat) print.printWhat = printWhat as PrintPropertiesOptions["printWhat"];
    const clrMode = attr(prnPr, "clrMode");
    if (clrMode && COLOR_MODE_FROM_XSD[clrMode]) print.colorMode = COLOR_MODE_FROM_XSD[clrMode];
    if (parseOnOff(prnPr.attributes?.["hiddenSlides"])) print.hiddenSlides = true;
    if (parseOnOff(prnPr.attributes?.["scaleToFitPaper"])) print.scaleToFitPaper = true;
    if (parseOnOff(prnPr.attributes?.["frameSlides"])) print.frameSlides = true;
    if (Object.keys(print).length > 0) result.print = print as PrintPropertiesOptions;
  }

  // htmlPublish (p:htmlPubPr)
  const htmlPubPr = findChild(el, "p:htmlPubPr");
  if (htmlPubPr) {
    const htmlPublish: NonNullable<PresentationPropertiesOptions["htmlPublish"]> = {};
    if (String(htmlPubPr.attributes?.["showSpeakerNotes"]) === "0") {
      htmlPublish.showSpeakerNotes = false;
    }
    const target = attr(htmlPubPr, "target");
    if (target) htmlPublish.target = target;
    const title = attr(htmlPubPr, "title");
    if (title) htmlPublish.title = title;
    const rId = attr(htmlPubPr, "r:id");
    if (rId) htmlPublish.rId = rId;
    if (Object.keys(htmlPublish).length > 0) result.htmlPublish = htmlPublish;
  }

  // colorMru (p:clrMru)
  const clrMru = findChild(el, "p:clrMru");
  if (clrMru) {
    const colors: string[] = [];
    for (const child of clrMru.elements ?? []) {
      if (child.name !== "a:srgbClr") continue;
      const val = attr(child, "val");
      if (val) colors.push(val);
    }
    if (colors.length > 0) result.colorMru = colors;
  }

  // p:extLst — verbatim inner XML for unmodeled extensions (p14 discardImageDpi etc.).
  const extLst = findChild(el, "p:extLst");
  if (extLst) {
    const inner = stringifyXml(extLst);
    if (inner) result.ext = inner;
  }

  return result as PresentationPropertiesOptions;
}
