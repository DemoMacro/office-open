import type { ParagraphChild, PictureOptions, ShapeOptions } from "@office-open/docx";

import type { RtfParseContext } from "../errors";
import { RtfParseError } from "../errors";
import type { RtfToken } from "../tokenizer";

const TWIPS_TO_EMU = 635;

export function imageDimensions(data: Uint8Array): { width: number; height: number } | undefined {
  if (data.length >= 24 && data[0] === 0x89 && data[1] === 0x50) {
    const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
    return { width: view.getUint32(16), height: view.getUint32(20) };
  }
  if (data.length >= 10 && data[0] === 0x47 && data[1] === 0x49) {
    return {
      width: data[6]! | (data[7]! << 8),
      height: data[8]! | (data[9]! << 8),
    };
  }
  if (data.length >= 4 && data[0] === 0xff && data[1] === 0xd8) {
    for (let index = 2; index < data.length - 8;) {
      if (data[index] !== 0xff) {
        index += 1;
        continue;
      }
      const marker = data[index + 1];
      if (
        marker === undefined ||
        marker === 0xd8 ||
        marker === 0x01 ||
        (marker >= 0xd0 && marker <= 0xd7)
      ) {
        index += 2;
        continue;
      }
      const length = (data[index + 2]! << 8) | data[index + 3]!;
      if (
        marker >= 0xc0 &&
        marker <= 0xcf &&
        marker !== 0xc4 &&
        marker !== 0xc8 &&
        marker !== 0xcc
      ) {
        return {
          height: (data[index + 5]! << 8) | data[index + 6]!,
          width: (data[index + 7]! << 8) | data[index + 8]!,
        };
      }
      index += 2 + length;
    }
  }
  return undefined;
}

export function parseRtfShape(
  tokens: readonly RtfToken[],
  start: number,
  end: number,
  _source: string,
): ParagraphChild {
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
  let picture: PictureOptions | undefined;

  const readPicture = (pictureStart: number, pictureEnd: number): PictureOptions => {
    const pictureContext: RtfParseContext = {
      part: "RTF picture destination",
      path: "\\pict",
      name: "picture",
      reason: "",
    };
    let type: "png" | "jpg" | "wmf" | "emf" = "png";
    let width = 0;
    let height = 0;
    let hex = "";
    const binary: number[] = [];
    for (let index = pictureStart; index < pictureEnd; index += 1) {
      const token = tokens[index];
      if (token?.kind === "control") {
        if (token.word === "pngblip") type = "png";
        else if (token.word === "jpegblip") type = "jpg";
        else if (token.word === "emfblip") type = "emf";
        else if (token.word === "wmetafile") type = "wmf";
        else if ((token.word === "picw" || token.word === "picwgoal") && token.param !== undefined)
          width = token.param * TWIPS_TO_EMU;
        else if ((token.word === "pich" || token.word === "pichgoal") && token.param !== undefined)
          height = token.param * TWIPS_TO_EMU;
      } else if (token?.kind === "hex") hex += token.value;
      else if (token?.kind === "binary") binary.push(...token.data);
      else if (token?.kind === "text") hex += token.value.replace(/[^0-9a-fA-F]/g, "");
    }
    const payload =
      binary.length > 0
        ? binary
        : (hex.match(/../gu)?.map((pair) => Number.parseInt(pair, 16)) ?? []);
    if (binary.length === 0 && hex.length % 2 !== 0) {
      pictureContext.reason = "odd hexadecimal byte length";
      throw new RtfParseError(
        "\\pict hexadecimal data has an odd byte length",
        0,
        "",
        pictureContext,
      );
    }
    const intrinsic = payload.length > 0 ? imageDimensions(new Uint8Array(payload)) : undefined;
    if (intrinsic) {
      width = width > 0 ? width : intrinsic.width * TWIPS_TO_EMU;
      height = height > 0 ? height : intrinsic.height * TWIPS_TO_EMU;
    }
    if (payload.length === 0 || width <= 0 || height <= 0) {
      pictureContext.reason = "missing data or dimensions";
      throw new RtfParseError(
        "\\pict requires binary data and positive dimensions",
        0,
        "",
        pictureContext,
      );
    }
    return {
      type,
      data: new Uint8Array(payload),
      transformation: { width, height },
    };
  };

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
        if (parameter < 0 || parameter > 5)
          throw shapeError("unsupported shape wrap type", context);
        if (parameter === 0) break;
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
        if (parameter === undefined || !wrap) break;
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
      case "chpleft":
      case "do":
      case "clvertalt":
      case "dobxcolumn":
      case "dobxmargin":
      case "dobxpage":
      case "pard":
      case "par":
      case "trrh":
      case "svb":
      case "trgaph":
      case "shprslt":
      case "trowd":
      case "cellx":
      case "cell":
      case "row":
      case "shpbxignore":
      case "shpbyignore":
      case "shplid":
      case "shp":
      case "shpgrp":
      case "shpinst":
      case "sp":
      case "sn":
      case "sv":
        break;
      case "pict":
        {
          let depth = 0;
          let cursor = index + 1;
          while (cursor < end) {
            const nested = tokens[cursor];
            if (nested?.kind === "group-start") depth += 1;
            else if (nested?.kind === "group-end") {
              if (depth === 0) break;
              depth -= 1;
            }
            cursor += 1;
          }
          picture = readPicture(index + 1, cursor);
          index = cursor - 1;
        }
        break;
      default:
        break;
    }
  }

  if (width === undefined && left !== undefined && right !== undefined) width = right - left;
  if (height === undefined && top !== undefined && bottom !== undefined) height = bottom - top;
  if (picture) return { picture };
  if (width === undefined || height === undefined) {
    throw shapeError("shape width and height are required", context);
  }

  return {
    wpsShape: {
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
    },
  };
}

function shapeError(reason: string, context: RtfParseContext): RtfParseError {
  context.reason = reason;
  return new RtfParseError(reason, 0, "", context);
}
