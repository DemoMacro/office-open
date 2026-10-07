import {
  attributeNumber,
  attributeString,
  childNamed,
  childrenNamed,
  textOf,
  xmlElement,
} from "@office-open/odf";
import type {
  AlignmentOptions,
  BorderSideOptions,
  FontOptions,
  StyleOptions,
} from "@office-open/xlsx";
import type { Element } from "@office-open/xml";

import { odfColor, parseBorder } from "./styles";

/** Translates an Excel numFmt code into the matching ODF data style element. */
export function numFmtDataStyle(numFmt: string): { name: string; children: string[] } {
  const percent = numFmt.endsWith("%");
  const core = percent ? numFmt.slice(0, -1) : numFmt;
  const tokens = core.match(/yyyy|yy|mm|dd|hh|ss|:/g);
  if (tokens && tokens.join("") === core) {
    const isTime = tokens.includes("hh") || tokens.includes("ss");
    if (isTime) {
      return {
        name: "number:time-style",
        children: tokens.map((token) => {
          const long = { hh: "number:hours", mm: "number:minutes", ss: "number:seconds" }[token];
          return long
            ? xmlElement(long, { "number:style": "long" })
            : xmlElement("number:text", undefined, [token]);
        }),
      };
    }
    return {
      name: "number:date-style",
      children: tokens.map((token) => {
        const element = {
          yyyy: "number:year",
          yy: "number:year",
          mm: "number:month",
          dd: "number:day",
        }[token];
        return element
          ? xmlElement(element, {
              "number:style": token.length === 2 && token !== "yy" ? "long" : undefined,
            })
          : xmlElement("number:text", undefined, [token]);
      }),
    };
  }
  const grouping = core.includes(",");
  const decimals = core.split(".")[1];
  if (!/^[#0,]*(\.[0]+)?$/.test(core)) {
    return {
      name: "number:number-style",
      children: [xmlElement("number:text", undefined, [numFmt])],
    };
  }
  const digits = Math.max(1, (core.split(".")[0] ?? "").replace(/[#,]/g, "").length);
  const children = [
    xmlElement("number:number", {
      "number:decimal-places": decimals?.length,
      "number:min-integer-digits": digits,
      "number:grouping": grouping || undefined,
    }),
  ];
  if (percent) children.push(xmlElement("number:text", undefined, ["%"]));
  return {
    name: percent ? "number:percentage-style" : "number:number-style",
    children,
  };
}

/** ODF office:date-value is UTC ISO 8601 without timezone suffix. */
export function odfDateValue(date: Date): string {
  return date.toISOString().slice(0, 19);
}

/** Maps table-cell style names back to typed StyleOptions via data styles. */
export function parseNumberStyles(container: Element | undefined): Map<string, StyleOptions> {
  const dataStyles = new Map<string, string>();
  const cellStyles = new Map<string, StyleOptions>();
  for (const element of container?.elements ?? []) {
    if (
      element.name === "number:number-style" ||
      element.name === "number:percentage-style" ||
      element.name === "number:date-style" ||
      element.name === "number:time-style"
    ) {
      const format = dataStyleNumFmt(element);
      if (format) dataStyles.set(attributeString(element, "style:name") ?? "", format);
    }
  }
  for (const style of childrenNamed(container, "style:style")) {
    if (attributeString(style, "style:family") !== "table-cell") continue;
    const options: StyleOptions = {};
    const dataName = attributeString(style, "style:data-style-name");
    const numFmt = dataName ? dataStyles.get(dataName) : undefined;
    if (numFmt) options.numFmt = numFmt;
    const text = childNamed(style, "style:text-properties");
    const size = attributeString(text, "fo:font-size");
    const fontColor = odfColor(attributeString(text, "fo:color"));
    const font: FontOptions = {
      color: fontColor ? { rgb: `FF${fontColor}` } : undefined,
      size: size?.endsWith("pt") ? Number(size.slice(0, -2)) : undefined,
      bold: attributeString(text, "fo:font-weight") === "bold" || undefined,
      italic: attributeString(text, "fo:font-style") === "italic" || undefined,
      underline: attributeString(text, "fo:underline-style") === "solid" || undefined,
      strike: attributeString(text, "fo:text-line-through-style") === "solid" || undefined,
    };
    if (Object.values(font).some((value) => value !== undefined)) options.font = font;
    const background = odfColor(
      attributeString(childNamed(style, "style:table-cell-properties"), "fo:background-color"),
    );
    if (background) options.fill = { type: "solid", foregroundColor: { rgb: `FF${background}` } };
    const cellProps = childNamed(style, "style:table-cell-properties");
    const border: BorderSideOptions = {
      top: parseBorder(attributeString(cellProps, "fo:border-top")),
      bottom: parseBorder(attributeString(cellProps, "fo:border-bottom")),
      left: parseBorder(attributeString(cellProps, "fo:border-left")),
      right: parseBorder(attributeString(cellProps, "fo:border-right")),
    };
    const diagonalDown = parseBorder(attributeString(cellProps, "style:diagonal-tl-br"));
    const diagonalUp = parseBorder(attributeString(cellProps, "style:diagonal-bl-tr"));
    if (diagonalDown !== undefined || diagonalUp !== undefined) {
      border.diagonal = diagonalDown ?? diagonalUp;
      border.diagonalDown = diagonalDown !== undefined || undefined;
      border.diagonalUp = diagonalUp !== undefined || undefined;
    }
    if (Object.values(border).some((side) => side !== undefined)) options.border = border;
    const vertical = attributeString(cellProps, "style:vertical-align");
    const wrapText = attributeString(cellProps, "fo:wrap-option") === "wrap";
    const horizontal = attributeString(
      childNamed(style, "style:paragraph-properties"),
      "fo:text-align",
    );
    const alignment: AlignmentOptions = {
      horizontal:
        horizontal === "start"
          ? "left"
          : horizontal === "end"
            ? "right"
            : (horizontal as AlignmentOptions["horizontal"]),
      vertical: (vertical === "middle" ? "center" : vertical) as AlignmentOptions["vertical"],
      wrapText: wrapText || undefined,
    };
    if (Object.values(alignment).some((value) => value !== undefined))
      options.alignment = alignment;
    if (Object.keys(options).length > 0) {
      cellStyles.set(attributeString(style, "style:name") ?? "", options);
    }
  }
  return cellStyles;
}

/** Reconstructs an Excel numFmt code from ODF number:* style children. */
export function dataStyleNumFmt(element: Element): string | undefined {
  const parts: string[] = [];
  for (const child of element.elements ?? []) {
    if (child.type === "text") {
      parts.push(String(child.text ?? ""));
      continue;
    }
    const long = attributeString(child, "number:style") === "long";
    switch (child.name) {
      case "number:number": {
        const decimals = attributeNumber(child, "number:decimal-places") ?? 0;
        const digits = attributeNumber(child, "number:min-integer-digits") ?? 1;
        parts.push(
          (attributeString(child, "number:grouping") === "true" ? "#,##0" : "0".repeat(digits)) +
            (decimals > 0 ? `.${"0".repeat(decimals)}` : ""),
        );
        break;
      }
      case "number:year":
        parts.push(long ? "yyyy" : "yy");
        break;
      case "number:month":
        parts.push(long ? "mm" : "m");
        break;
      case "number:day":
        parts.push(long ? "dd" : "d");
        break;
      case "number:hours":
        parts.push(long ? "hh" : "h");
        break;
      case "number:minutes":
        parts.push(long ? "mm" : "m");
        break;
      case "number:seconds":
        parts.push(long ? "ss" : "s");
        break;
      case "number:text":
        parts.push(textOf(child));
        break;
      default:
        return undefined;
    }
  }
  let format = parts.join("");
  if (element.name === "number:percentage-style" && !format.endsWith("%")) format += "%";
  return format;
}
