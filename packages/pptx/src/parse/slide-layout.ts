import { attr, findChild } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type { SlideLayoutType } from "../file/slide-layout/slide-layout";

// Display name → SlideLayoutType mapping (from LAYOUT_DEFS in slide-layout.ts)
const NAME_TO_TYPE: Record<string, SlideLayoutType> = {
  "Title Slide": "title",
  "Title and Content": "obj",
  "Section Header": "secHead",
  "Two Content": "twoObj",
  Comparison: "twoTxTwoObj",
  "Title Only": "titleOnly",
  Blank: "blank",
  "Content with Caption": "objTx",
  "Picture with Caption": "picTx",
  "Vertical Text": "vertTx",
  "Vertical Title and Text": "vertTitleAndTx",
  "Title and Text": "tx",
};

export function parseSlideLayoutType(el: Element): SlideLayoutType | string {
  const cSld = findChild(el, "p:cSld");
  if (cSld) {
    const name = attr(cSld, "name");
    if (name) {
      const mapped = NAME_TO_TYPE[name];
      if (mapped) return mapped;
      // Fallback: return display name as custom string
      return name;
    }
  }
  return "blank";
}
