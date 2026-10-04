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
  let width: number | undefined;
  let height: number | undefined;
  let text = "";
  let hasText = false;

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
      case "shpwidth":
        if (parameter === undefined) throw shapeError("extent requires a value", context);
        width = parameter * TWIPS_TO_EMU;
        break;
      case "shpheight":
        if (parameter === undefined) throw shapeError("extent requires a value", context);
        height = parameter * TWIPS_TO_EMU;
        break;
      case "shptxt":
        hasText = true;
        break;
      default:
        context.reason = `unsupported control \\${token.word}`;
        throw new RtfParseError(context.reason, token.position, source, context);
    }
  }

  if (hasText) {
    for (let index = start; index < end; index += 1) {
      const token = tokens[index];
      if (token?.kind === "text") text += token.value;
      if (token?.kind === "hex") text += String.fromCharCode(Number.parseInt(token.value, 16));
    }
  }
  if (width === undefined || height === undefined) {
    throw shapeError("shape width and height are required", context);
  }

  return {
    children: text ? [text] : [],
    transformation: { width, height },
    ...(left !== undefined || top !== undefined
      ? {
          floating: {
            horizontalPosition: {
              relative: "column",
              ...(left !== undefined ? { offset: left } : {}),
            },
            verticalPosition: {
              relative: "paragraph",
              ...(top !== undefined ? { offset: top } : {}),
            },
          },
        }
      : {}),
  };
}

function shapeError(reason: string, context: RtfParseContext): RtfParseError {
  context.reason = reason;
  return new RtfParseError(reason, 0, "", context);
}
