import type { RtfGenerateContext } from "./context";
import { control, rtfText } from "./escape";

export function writeFontTable(context: RtfGenerateContext): string {
  if (context.fonts.length === 0) return "";
  const entries = context.fonts
    .map((font, index) => `\\f${index}\\fswiss ${rtfText(font)};`)
    .join("");
  return `{\\fonttbl${entries}}`;
}

export function writeColorTable(context: RtfGenerateContext): string {
  if (context.colors.length === 0) return "";
  const entries = context.colors
    .map((color) => {
      const red = Number.parseInt(color.slice(0, 2), 16);
      const green = Number.parseInt(color.slice(2, 4), 16);
      const blue = Number.parseInt(color.slice(4, 6), 16);
      return `\\red${red}\\green${green}\\blue${blue};`;
    })
    .join("");
  return `{\\colortbl;${entries}}`;
}

export { control };
