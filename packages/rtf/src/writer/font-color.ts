import type { RtfGenerateContext } from "./context";
import { reject } from "./context";
import { rtfText } from "./escape";

const FAMILY_CONTROLS = {
  roman: "\\froman",
  swiss: "\\fswiss",
  modern: "\\fmodern",
  decor: "\\fdecor",
  script: "\\fscript",
  tech: "\\ftech",
  bidi: "\\fbidi",
} as const;

type FontFamily = keyof typeof FAMILY_CONTROLS;

export function writeFontTable(context: RtfGenerateContext): string {
  if (context.fonts.length === 0) return "";
  const entries = context.fonts
    .map((font, index) => {
      const familyControl =
        font.family === undefined ? undefined : FAMILY_CONTROLS[font.family as FontFamily];
      if (font.family !== undefined && familyControl === undefined) {
        reject("font-table", `fonts[${index}]`, "family", "RTF font family is unknown");
      }
      const family = familyControl ?? "";
      const charset =
        font.characterSet === undefined ? undefined : Number.parseInt(font.characterSet, 16);
      if (charset !== undefined && (!Number.isInteger(charset) || charset < 0 || charset > 255)) {
        reject("font-table", `fonts[${index}]`, "characterSet", "RTF charset requires 0-255");
      }
      const controls = `${family}${charset === undefined ? "" : `\\fcharset${charset}`}`;
      return `{\\f${index}${controls} ${rtfText(font.name ?? "")};}`;
    })
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
