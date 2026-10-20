/**
 * Shared cSld sub-block serializers/parsers (CT_CommonSlideData tail children)
 * used by the slide-family serializers — compiler.ts (slide), slide-master.ts,
 * slide-layout.ts, and the slide descriptor. p:hf is attribute-form per
 * CT_HeaderFooter (sldNum/hdr/ftr/dt boolean attributes, no element children);
 * parse additionally tolerates the legacy element-form emission.
 *
 * @module
 */

import { parseOnOff } from "@office-open/core";
import { attr, attrNum, escapeXml, findChild } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import type { SlideCustomerDataReferenceOptions } from "@shared/customer-data";
import type { SlideHeaderFooterOptions } from "@shared/header-footer";
import type { PictureOptions } from "@shared/picture";

import type { ControlOptions } from "./slide";

// ── p:custDataLst ──

/** Serialize cSld customer-data references; empty string when there are none. */
export function stringifyCustDataLst(
  items: SlideCustomerDataReferenceOptions[] | undefined,
): string {
  if (!items || items.length === 0) return "";
  return `<p:custDataLst>${items
    .map((item) =>
      item.kind === "tags" ? `<p:tags r:id="${item.rId}"/>` : `<p:custData r:id="${item.rId}"/>`,
    )
    .join("")}</p:custDataLst>`;
}

/** Parse a p:custDataLst element into customer-data references. */
export function parseCustDataLst(
  el: Element | undefined,
): SlideCustomerDataReferenceOptions[] | undefined {
  if (!el) return undefined;
  const items: SlideCustomerDataReferenceOptions[] = [];
  for (const cd of el.elements ?? []) {
    if (cd.name === "p:tags") {
      const rId = attr(cd, "r:id");
      if (rId) items.push({ rId, kind: "tags" });
    } else if (cd.name === "p:custData") {
      const rId = attr(cd, "r:id");
      if (rId) items.push({ rId });
    }
  }
  return items.length > 0 ? items : undefined;
}

// ── p:controls ──

/** Serialize cSld controls; empty string when there are none. */
export function stringifyControls(
  controls: ControlOptions[] | undefined,
  stringifyPicture?: (picture: PictureOptions) => string | undefined,
): string {
  if (!controls || controls.length === 0) return "";
  const items = controls
    .map((control) => {
      if (control.alternateContent) {
        const alternate = control.alternateContent;
        const namespace =
          alternate.requiresPrefix && alternate.requiresNamespace
            ? ` xmlns:${alternate.requiresPrefix}="${escapeXml(alternate.requiresNamespace)}"`
            : "";
        const requires = alternate.requires ? ` Requires="${escapeXml(alternate.requires)}"` : "";
        return (
          '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
          `<mc:Choice${namespace}${requires}>${stringifyControlElement(
            alternate.choice,
            stringifyPicture,
          )}</mc:Choice>` +
          `<mc:Fallback>${stringifyControlElement(
            alternate.fallback,
            stringifyPicture,
          )}</mc:Fallback></mc:AlternateContent>`
        );
      }
      return stringifyControlElement(control, stringifyPicture);
    })
    .join("");
  return `<p:controls>${items}</p:controls>`;
}

function stringifyControlElement(
  control: ControlOptions,
  stringifyPicture?: (picture: PictureOptions) => string | undefined,
): string {
  const attrs: string[] = [];
  if (control.shapeId !== undefined) attrs.push(`spid="${escapeXml(control.shapeId)}"`);
  if (control.name) attrs.push(`name="${escapeXml(control.name)}"`);
  if (control.showAsIcon) attrs.push('showAsIcon="1"');
  if (control.rId) attrs.push(`r:id="${control.rId}"`);
  if (control.imageWidth !== undefined) attrs.push(`imgW="${control.imageWidth}"`);
  if (control.imageHeight !== undefined) attrs.push(`imgH="${control.imageHeight}"`);
  const picture = control.picture && stringifyPicture ? stringifyPicture(control.picture) : "";
  const opening = `<p:control${attrs.length > 0 ? ` ${attrs.join(" ")}` : ""}`;
  return picture ? `${opening}>${picture}</p:control>` : `${opening}/>`;
}

/** Parse a p:controls element into control options. */
export function parseControls(
  el: Element | undefined,
  parsePicture?: (picture: Element) => PictureOptions | undefined,
): ControlOptions[] | undefined {
  if (!el) return undefined;
  const items: ControlOptions[] = [];
  for (const ctrl of el.elements ?? []) {
    if (ctrl.name === "mc:AlternateContent") {
      const choiceEl = findChild(ctrl, "mc:Choice");
      const fallbackEl = findChild(ctrl, "mc:Fallback");
      const choiceControl = choiceEl ? findChild(choiceEl, "p:control") : undefined;
      const fallbackControl = fallbackEl ? findChild(fallbackEl, "p:control") : undefined;
      if (!choiceControl || !fallbackControl) continue;
      const requires = attr(choiceEl, "Requires") ?? undefined;
      const requiresPrefix = requires ?? undefined;
      const requiresNamespace = requires ? attr(choiceEl, `xmlns:${requires}`) : undefined;
      items.push({
        alternateContent: {
          ...(requires ? { requires } : {}),
          ...(requiresPrefix ? { requiresPrefix } : {}),
          ...(requiresNamespace ? { requiresNamespace } : {}),
          choice: parseControlElement(choiceControl, parsePicture),
          fallback: parseControlElement(fallbackControl, parsePicture),
        },
      });
    } else if (ctrl.name === "p:control") {
      items.push(parseControlElement(ctrl, parsePicture));
    }
  }
  return items.length > 0 ? items : undefined;
}

function parseControlElement(
  ctrl: Element,
  parsePicture?: (picture: Element) => PictureOptions | undefined,
): ControlOptions {
  const item: ControlOptions = {};
  const shapeId = attr(ctrl, "spid");
  if (shapeId) item.shapeId = shapeId;
  const name = attr(ctrl, "name");
  if (name) item.name = name;
  if (parseOnOff(attr(ctrl, "showAsIcon"))) item.showAsIcon = true;
  const rId = attr(ctrl, "r:id");
  if (rId) item.rId = rId;
  const imageWidth = attrNum(ctrl, "imgW");
  if (imageWidth !== undefined) item.imageWidth = imageWidth;
  const imageHeight = attrNum(ctrl, "imgH");
  if (imageHeight !== undefined) item.imageHeight = imageHeight;
  const pictureEl = findChild(ctrl, "p:pic");
  const picture = pictureEl && parsePicture ? parsePicture(pictureEl) : undefined;
  if (picture) item.picture = picture;
  return item;
}

// ── p:hf (CT_HeaderFooter, attribute form) ──

/**
 * Serialize a layout-level p:hf in attribute form per CT_HeaderFooter;
 * undefined/empty options emit nothing (each attribute defaults to true in
 * the XSD, so an all-true header/footer is the absent element).
 */
export function stringifySlideHf(hf: SlideHeaderFooterOptions | undefined): string {
  if (!hf) return "";
  const attrs: string[] = [];
  if (hf.slideNumber !== undefined) attrs.push(`sldNum="${hf.slideNumber ? 1 : 0}"`);
  if (hf.dateTime !== undefined) attrs.push(`dt="${hf.dateTime ? 1 : 0}"`);
  if (hf.footer !== undefined) attrs.push(`ftr="${hf.footer ? 1 : 0}"`);
  if (hf.header !== undefined) attrs.push(`hdr="${hf.header ? 1 : 0}"`);
  if (attrs.length === 0) return "";
  return `<p:hf ${attrs.join(" ")}/>`;
}

/**
 * Parse a layout-level p:hf. Reads CT_HeaderFooter attributes; tolerates the
 * legacy element-form children this library used to emit.
 */
export function parseSlideHf(el: Element | undefined): SlideHeaderFooterOptions | undefined {
  if (!el) return undefined;
  const hf: SlideHeaderFooterOptions = {};
  const sldNum = attr(el, "sldNum");
  if (sldNum !== undefined) hf.slideNumber = parseOnOff(sldNum) ?? false;
  const dt = attr(el, "dt");
  if (dt !== undefined) hf.dateTime = parseOnOff(dt) ?? false;
  const ftr = attr(el, "ftr");
  if (ftr !== undefined) hf.footer = parseOnOff(ftr) ?? false;
  const hdr = attr(el, "hdr");
  if (hdr !== undefined) hf.header = parseOnOff(hdr) ?? false;
  if (Object.keys(hf).length > 0) return hf;

  // Legacy element-form emission (<p:hf><p:sldNum/>…</p:hf>).
  if (findChild(el, "p:sldNum")) hf.slideNumber = true;
  if (findChild(el, "p:dt")) hf.dateTime = true;
  if (findChild(el, "p:ftr")) hf.footer = true;
  if (findChild(el, "p:hdr")) hf.header = true;
  return Object.keys(hf).length > 0 ? hf : undefined;
}
