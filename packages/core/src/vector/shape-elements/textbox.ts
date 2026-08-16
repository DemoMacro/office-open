/**
 * v:textbox element — CT_Textbox.
 *
 * The content model is a choice: w:txbxContent (WordprocessingML paragraphs,
 * owned by the docx package) or an unqualified local element (the HTML `<div>`
 * Excel writes). This domain stays format-agnostic by carrying both as raw
 * inner-XML strings the caller produces/consumes.
 *
 * Reference: ISO/IEC 29500-4, vml-main.xsd, CT_Textbox.
 *
 * @module
 */
import type { Element as XmlElement } from "@office-open/xml";
import { escapeXml, stringifyElement } from "@office-open/xml";

import { stringifyVmlTrueFalse, parseVmlTrueFalse } from "../attributes";
import { parseVmlShapeStyle, parseVmlStyle, stringifyVmlStyle, type VmlShapeStyle } from "../style";
import type { VmlInsetMode } from "./office-elements";

/** v:textbox options (CT_Textbox). */
export interface VmlTextboxOptions {
  id?: string;
  style?: VmlShapeStyle;
  /** Inner inset, e.g. "auto" or "10pt,10pt,10pt,10pt". */
  inset?: string;
  /** o:singleclick — the whole shape is the click target. */
  singleclick?: boolean;
  /** o:insetmode — inset handling mode. */
  insetmode?: VmlInsetMode;
  /** Inner XML of the w:txbxContent child (docx paragraphs), caller-produced. */
  txbxContent?: string;
  /** Verbatim inner XML for a local (unqualified) child — Excel's `<div>` form. */
  content?: string;
}

/**
 * Serialize v:textbox. Exactly one content form is emitted when present
 * (txbxContent wins if both are set, matching the XSD choice).
 */
export function stringifyVmlTextbox(opts: VmlTextboxOptions): string {
  const attrs: string[] = [];
  if (opts.id !== undefined) attrs.push(`id="${escapeXml(opts.id)}"`);
  if (opts.style !== undefined) attrs.push(`style="${escapeXml(stringifyVmlStyle(opts.style))}"`);
  if (opts.inset !== undefined) attrs.push(`inset="${escapeXml(opts.inset)}"`);
  if (opts.singleclick !== undefined) {
    attrs.push(`o:singleclick="${stringifyVmlTrueFalse(opts.singleclick)}"`);
  }
  if (opts.insetmode !== undefined) attrs.push(`o:insetmode="${opts.insetmode}"`);
  const attrStr = attrs.length > 0 ? ` ${attrs.join(" ")}` : "";

  if (opts.txbxContent !== undefined) {
    return `<v:textbox${attrStr}><w:txbxContent>${opts.txbxContent}</w:txbxContent></v:textbox>`;
  }
  if (opts.content !== undefined) {
    return `<v:textbox${attrStr}>${opts.content}</v:textbox>`;
  }
  return `<v:textbox${attrStr}/>`;
}

/** Parse a v:textbox element. */
export function parseVmlTextbox(el: XmlElement): VmlTextboxOptions {
  const out: VmlTextboxOptions = {};
  const attrs = el.attributes ?? {};
  if (attrs.id !== undefined) out.id = String(attrs.id);
  if (attrs.style !== undefined) {
    out.style = parseVmlShapeStyle(parseVmlStyle(String(attrs.style)));
  }
  if (attrs.inset !== undefined) out.inset = String(attrs.inset);
  if (attrs["o:singleclick"] !== undefined) {
    out.singleclick = parseVmlTrueFalse(String(attrs["o:singleclick"]));
  }
  if (attrs["o:insetmode"] !== undefined) {
    out.insetmode = String(attrs["o:insetmode"]) as VmlInsetMode;
  }

  for (const child of el.elements ?? []) {
    if (child.type !== "element") continue;
    if (child.name === "w:txbxContent") {
      out.txbxContent = (child.elements ?? []).map((e) => stringifyElement(e)).join("");
      break;
    }
    // First local (unqualified) child wins — the XSD is a single-element choice.
    if (!child.name?.includes(":")) {
      out.content = stringifyElement(child);
      break;
    }
  }
  return out;
}
