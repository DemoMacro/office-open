import type { ShapeOptions } from "@office-open/docx";

import type { RtfParseContext } from "../errors";
import { RtfParseError } from "../errors";
import type { RtfToken } from "../tokenizer";

const TWIPS_TO_EMU = 635;

export function parseRtfShape(
  tokens: readonly RtfToken[],
  start: number,
  end: number,
  source: string,
): ShapeOptions {
  const context: RtfParseContext = {
    part: "RTF shape destination",
    path: "\\shpinst",
    name: "shape",
    reason: "",
  };
  let left: number | undefined;
  let top: number | undefined;
  let right: number | undefined;
  let bottom: number | undefined;
  let width: number | undefined;
  let height: number | undefined;
  let horizontalRelative: NonNullable<
    NonNullable<ShapeOptions["floating"]>["horizontalPosition"]
  >["relative"] = "column";
  let verticalRelative: NonNullable<
    NonNullable<ShapeOptions["floating"]>["verticalPosition"]
  >["relative"] = "paragraph";
  let wrap: NonNullable<NonNullable<ShapeOptions["floating"]>["wrap"]> | undefined;
  let zIndex: number | undefined;
  let behindDocument = false;
  let lockAnchor = false;
  let text = "";

  const decodeShapeText = (start: number, end: number): string => {
    let value = "";
    let unicodeSkip = 0;
    let unicodeByteCount = 1;
    for (let index = start; index < end; index += 1) {
      const token = tokens[index];
      if (!token) continue;
      if (unicodeSkip > 0 && (token.kind === "text" || token.kind === "hex")) {
        unicodeSkip -= 1;
        continue;
      }
      if (token.kind === "text") value += token.value;
      else if (token.kind === "hex") value += String.fromCharCode(Number.parseInt(token.value, 16));
      else if (token.kind === "control") {
        if (token.word === "par" || token.word === "line") value += "\n";
        else if (token.word === "tab") value += "\t";
        else if (token.word === "uc") unicodeByteCount = token.param ?? 1;
        else if (token.word === "u" && token.param !== undefined) {
          const codePoint = token.param < 0 ? token.param + 65536 : token.param;
          value += String.fromCodePoint(codePoint);
          unicodeSkip = unicodeByteCount;
        } else if (token.word === undefined) {
          if (token.symbol === "~") value += " ";
          else if (token.symbol === "_") value += "‐";
          else if (token.symbol === "-" || token.symbol === "\r" || token.symbol === "\n") continue;
          else if (token.symbol) value += token.symbol;
        }
      }
    }
    return value;
  };

  for (let index = start; index < end; index += 1) {
    const token = tokens[index];
    if (token?.kind !== "control") continue;
    if (token.word === undefined) continue;
    const parameter = token.param;
    switch (token.word) {
      case "shpleft":
        if (parameter === undefined) throw shapeError("position requires a value", context);
        left = parameter * TWIPS_TO_EMU;
        break;
      case "shptop":
        if (parameter === undefined) throw shapeError("position requires a value", context);
        top = parameter * TWIPS_TO_EMU;
        break;
      case "shpright":
        if (parameter === undefined) throw shapeError("position requires a value", context);
        right = parameter * TWIPS_TO_EMU;
        break;
      case "shpbottom":
        if (parameter === undefined) throw shapeError("position requires a value", context);
        bottom = parameter * TWIPS_TO_EMU;
        break;
      case "shpwidth":
        if (parameter === undefined) throw shapeError("extent requires a value", context);
        width = parameter * TWIPS_TO_EMU;
        break;
      case "shpheight":
        if (parameter === undefined) throw shapeError("extent requires a value", context);
        height = parameter * TWIPS_TO_EMU;
        break;
      case "shpwr":
        if (parameter === undefined) throw shapeError("wrap type requires a value", context);
        if (parameter < 1 || parameter > 5)
          throw shapeError("unsupported shape wrap type", context);
        wrap = {
          type:
            parameter === 1
              ? "topAndBottom"
              : parameter === 2
                ? "square"
                : parameter === 3
                  ? "none"
                  : parameter === 4
                    ? "tight"
                    : "through",
        };
        break;
      case "shpwrk":
        if (parameter === undefined || !wrap) throw shapeError("unexpected wrap side", context);
        if (parameter > 3) throw shapeError("unsupported shape wrap side", context);
        wrap.side =
          parameter === 0
            ? "bothSides"
            : parameter === 1
              ? "left"
              : parameter === 2
                ? "right"
                : "largest";
        break;
      case "shpfblwtxt":
        if (parameter === undefined) throw shapeError("z-order requires a value", context);
        behindDocument = parameter !== 0;
        break;
      case "shpz":
        if (parameter === undefined) throw shapeError("z-order requires a value", context);
        zIndex = parameter;
        break;
      case "shplockanchor":
        lockAnchor = true;
        break;
      case "shpbxpage":
        horizontalRelative = "page";
        break;
      case "shpbxmargin":
        horizontalRelative = "margin";
        break;
      case "shpbxcolumn":
        horizontalRelative = "column";
        break;
      case "shpbypage":
        verticalRelative = "page";
        break;
      case "shpbymargin":
        verticalRelative = "margin";
        break;
      case "shpbypara":
        verticalRelative = "paragraph";
        break;
      case "shptxt":
        if (tokens[index + 1]?.kind === "group-start") {
          let depth = 1;
          let cursor = index + 2;
          while (cursor < end && depth > 0) {
            const nested = tokens[cursor];
            if (nested?.kind === "group-start") depth += 1;
            else if (nested?.kind === "group-end") depth -= 1;
            cursor += 1;
          }
          const bodyEnd = depth === 0 ? cursor - 1 : end;
          text += decodeShapeText(index + 2, bodyEnd);
          index = bodyEnd - 1;
        } else {
          text += decodeShapeText(index + 1, end);
          index = end;
        }
        break;
      case "shpfhdr":
      case "shpbxignore":
      case "shpbyignore":
      case "shplid":
      case "sp":
      case "sn":
      case "sv":
        break;
      default:
        context.reason = `unsupported control \\${token.word}`;
        throw new RtfParseError(context.reason, token.position, source, context);
    }
  }

  if (width === undefined && left !== undefined && right !== undefined) width = right - left;
  if (height === undefined && top !== undefined && bottom !== undefined) height = bottom - top;
  if (width === undefined || height === undefined) {
    throw shapeError("shape width and height are required", context);
  }

  return {
    children: text ? [text] : [],
    transformation: { width, height },
    ...(left !== undefined ||
    top !== undefined ||
    zIndex !== undefined ||
    wrap ||
    behindDocument ||
    lockAnchor
      ? {
          floating: {
            horizontalPosition: {
              relative: horizontalRelative,
              ...(left !== undefined ? { offset: left } : {}),
            },
            verticalPosition: {
              relative: verticalRelative,
              ...(top !== undefined ? { offset: top } : {}),
            },
            ...(behindDocument ? { behindDocument } : {}),
            ...(lockAnchor ? { lockAnchor } : {}),
            ...(zIndex !== undefined ? { zIndex } : {}),
            ...(wrap ? { wrap } : {}),
          },
        }
      : {}),
  };
}

function shapeError(reason: string, context: RtfParseContext): RtfParseError {
  context.reason = reason;
  return new RtfParseError(reason, 0, "", context);
}
