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

import { OdsParseError } from "./error";
import { odfColor, parseBorder, parseOdfStrike, parseOdfUnderline } from "./styles";

/** Translates an Excel numFmt code into the matching ODF data style element. */
export function numFmtDataStyle(numFmt: string): { name: string; children: string[] } {
  const currency = /^\[\$([^-\]]+)(?:-[^\]]+)?\]/u.exec(numFmt);
  if (currency) {
    const symbol = currency[1]!;
    const core = numFmt.slice(currency[0].length);
    const decimals = core.split(".")[1];
    const digits = Math.max(1, (core.split(".")[0] ?? "").replace(/[#,]/g, "").length);
    return {
      name: "number:currency-style",
      children: [
        xmlElement("number:currency-symbol", undefined, [symbol]),
        xmlElement("number:number", {
          "number:decimal-places": decimals?.length,
          "number:min-integer-digits": digits,
          "number:grouping": core.includes(",") || undefined,
        }),
      ],
    };
  }
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
  return date.getUTCMilliseconds()
    ? date.toISOString().slice(0, 23)
    : date.toISOString().slice(0, 19);
}

/** Maps table-cell style names back to typed StyleOptions via data styles. */
export function parseNumberStyles(container: Element | undefined): Map<string, StyleOptions> {
  const dataStyles = new Map<string, string>();
  const cellStyles = new Map<string, StyleOptions>();
  const stylesByName = new Map(
    (container?.elements ?? [])
      .filter((element) => element.name?.startsWith("number:"))
      .map((element) => [attributeString(element, "style:name") ?? "", element]),
  );
  for (const element of container?.elements ?? []) {
    if (
      element.name === "number:number-style" ||
      element.name === "number:percentage-style" ||
      element.name === "number:currency-style" ||
      element.name === "number:date-style" ||
      element.name === "number:time-style"
    ) {
      const format = dataStyleNumFmt(element, {
        stylesByName,
        visiting: new Set(),
      });
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
      ...parseOdfUnderline(text),
      ...parseOdfStrike(text),
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
    const verticalValue = attributeString(cellProps, "style:vertical-align");
    const vertical =
      verticalValue === "automatic" || verticalValue === "auto" ? undefined : verticalValue;
    const wrapText = attributeString(cellProps, "fo:wrap-option") === "wrap";
    const horizontal = attributeString(
      childNamed(style, "style:paragraph-properties"),
      "fo:text-align",
    );
    const alignment: AlignmentOptions = {
      horizontal:
        horizontal === undefined
          ? undefined
          : horizontal === "start"
            ? "left"
            : horizontal === "end"
              ? "right"
              : (["left", "center", "right", "justify"] as const).find(
                  (value) => value === horizontal,
                ),
      vertical:
        vertical === undefined
          ? undefined
          : vertical === "middle"
            ? "center"
            : (["top", "bottom"] as const).find((value) => value === vertical),
      wrapText: wrapText || undefined,
    };
    if (
      (horizontal !== undefined && !alignment.horizontal) ||
      (vertical !== undefined && !alignment.vertical)
    )
      throw new OdsParseError(
        "Invalid OpenDocument cell alignment",
        "content.xml",
        "/office:document-content/office:automatic-styles/style:style",
        vertical !== undefined && !alignment.vertical ? "style:vertical-align" : "fo:text-align",
        "unsupported alignment value",
      );
    if (Object.values(alignment).some((value) => value !== undefined))
      options.alignment = alignment;
    if (Object.keys(options).length > 0) {
      cellStyles.set(attributeString(style, "style:name") ?? "", options);
    }
  }
  return cellStyles;
}

/** Reconstructs an Excel numFmt code from ODF number:* style children. */
export function dataStyleNumFmt(
  element: Element,
  styleMap: {
    stylesByName: Map<string, Element>;
    visiting: Set<string>;
  } = { stylesByName: new Map(), visiting: new Set() },
): string | undefined {
  const styleName = attributeString(element, "style:name");
  if (styleName && styleMap.visiting.has(styleName)) {
    throw new OdsParseError(
      "content.xml: cyclic number style reference",
      "content.xml",
      "/office:automatic-styles",
      styleName,
      "cyclic-style-reference",
    );
  }
  if (styleName) styleMap.visiting.add(styleName);
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
      case "number:currency-symbol":
        parts.push(`[$${textOf(child)}]`);
        break;
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
      case "loext:fill-character":
      case "number:fill-character":
        parts.push(`*${textOf(child)}`);
        break;
      case "loext:text":
        parts.push(textOf(child));
        break;
      case "number:text-content":
        parts.push("@");
        break;
      case "number:fraction": {
        const numerator =
          attributeNumber(child, "loext:max-numerator-digits") ??
          attributeNumber(child, "number:min-numerator-digits") ??
          1;
        const denominatorValue =
          attributeNumber(child, "loext:max-denominator-value") ??
          attributeNumber(child, "number:max-denominator-value");
        const denominator = denominatorValue
          ? String(denominatorValue).length
          : (attributeNumber(child, "number:min-denominator-digits") ?? 1);
        parts.push(`# ${"?".repeat(numerator)}/${"?".repeat(denominator)}`);
        break;
      }
      case "style:text-properties": {
        const color = attributeString(child, "fo:color");
        const colorName = numberFormatColor(color);
        if (colorName) parts.push(colorName);
        break;
      }
      case "style:map": {
        const target = styleMap.stylesByName.get(
          attributeString(child, "style:apply-style-name") ?? "",
        );
        if (!target) {
          throw new OdsParseError(
            `content.xml: missing number style "${attributeString(child, "style:apply-style-name") ?? ""}"`,
            "content.xml",
            "/office:automatic-styles",
            child.name,
            "missing-number-style",
          );
        }
        const mapped = dataStyleNumFmt(target, styleMap);
        if (mapped) parts.push(mapped);
        break;
      }
      case "number:scientific-number": {
        const mantissa = attributeNumber(child, "number:mantissa-places") ?? 1;
        const exponent = attributeNumber(child, "number:min-exponent-digits") ?? 2;
        parts.push(`0.${"0".repeat(mantissa)}E+${"0".repeat(exponent)}`);
        break;
      }
      case "number:am-pm":
        parts.push("AM/PM");
        break;
      case "number:boolean":
        parts.push("TRUE;FALSE");
        break;
      case "number:day-of-week":
        parts.push(long ? "dddd" : "ddd");
        break;
      case "number:quarter":
        parts.push(long ? "qq" : "q");
        break;
      case "number:era":
        parts.push(long ? "ggg" : "gg");
        break;
      case "number:week-of-year":
        parts.push("ww");
        break;
      default:
        throw new OdsParseError(
          `content.xml: unsupported number format element "${child.name}"`,
          "content.xml",
          "/office:automatic-styles",
          child.name,
          "unsupported-number-format-element",
        );
    }
  }
  if (styleName) styleMap.visiting.delete(styleName);
  let format = parts.join("");
  if (element.name === "number:percentage-style" && !format.endsWith("%")) format += "%";
  return format;
}

const NUMBER_FORMAT_COLORS: Record<string, string> = {
  "#000000": "[Black]",
  "#0000ff": "[Blue]",
  "#00ffff": "[Cyan]",
  "#00ff00": "[Green]",
  "#ff00ff": "[Magenta]",
  "#ff0000": "[Red]",
  "#ffffff": "[White]",
  "#ffff00": "[Yellow]",
};

function numberFormatColor(color: string | undefined): string | undefined {
  return color === undefined ? undefined : NUMBER_FORMAT_COLORS[color.toLowerCase()];
}
