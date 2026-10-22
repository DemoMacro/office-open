import type { RunOptions } from "@office-open/docx";

import { tagToLcid } from "../langs";
import {
  assertAllowed,
  colorIndex,
  fontIndex,
  reject,
  styleCode,
  type RtfGenerateContext,
} from "./context";
import { control, rtfText } from "./escape";
import { writeFieldChild } from "./field";
import { writeEndnoteReference } from "./footnote";
import { writeDrawingChild } from "./shape-picture";

const RUN_PROPERTIES = [
  "additionRsid",
  "bold",
  "italic",
  "underline",
  "color",
  "size",
  "font",
  "strike",
  "verticalAlign",
  "language",
  "characterSpacing",
  "kern",
  "allCaps",
  "smallCaps",
  "outline",
  "shadow",
  "vanish",
  "emboss",
  "scale",
  "style",
  "highlight",
] as const;

const RTF_HIGHLIGHT_INDICES: Record<string, number> = {
  black: 1,
  blue: 2,
  cyan: 3,
  green: 4,
  magenta: 5,
  red: 6,
  yellow: 7,
  darkBlue: 9,
  darkCyan: 10,
  darkGreen: 11,
  darkMagenta: 12,
  darkRed: 13,
  darkYellow: 14,
  darkGray: 15,
  lightGray: 16,
};

const UNDERLINE_CONTROLS: Record<string, string> = {
  single: "ul",
  words: "ulw",
  dotted: "uldb",
  dash: "uldash",
  dotDash: "uldashd",
  dotDotDash: "uldashdd",
  thick: "ulth",
  wave: "ulwave",
  wavyHeavy: "ulhwave",
  dashLong: "ulldash",
};

export function writeRunProperties(run: RunOptions, context: RtfGenerateContext): string {
  assertAllowed(run, [...RUN_PROPERTIES, "text", "children"], "run", "run");
  return [
    run.bold === true ? control("b") : "",
    run.bold === false ? control("b", 0) : "",
    run.additionRsid === undefined
      ? ""
      : control(
          "insrsid",
          Number.parseInt(run.additionRsid, 16) ||
            reject(
              "run",
              "run.additionRsid",
              run.additionRsid,
              "RTF writer requires hexadecimal revision save IDs",
            ),
        ),
    run.italic === true ? control("i") : "",
    run.italic === false ? control("i", 0) : "",
    run.strike === true ? control("strike") : "",
    run.strike === false ? control("strike", 0) : "",
    (() => {
      if (!run.underline) return "";
      const style = typeof run.underline === "object" ? run.underline.type : undefined;
      const controlName = style !== undefined ? UNDERLINE_CONTROLS[style] : "ul";
      if (controlName === undefined)
        return reject(
          "run",
          "run.underline.type",
          String(style),
          "unsupported RTF underline style",
        );
      return control(controlName);
    })(),
    (() => {
      if (run.highlight === undefined || run.highlight === "none") return "";
      const index = RTF_HIGHLIGHT_INDICES[run.highlight];
      if (index === undefined)
        return reject("run", "run.highlight", run.highlight, "unsupported RTF highlight color");
      return control("highlight", index);
    })(),
    run.size === undefined
      ? ""
      : typeof run.size === "number"
        ? control("fs", Math.round(run.size * 2))
        : reject("run", "run.size", String(run.size), "RTF writer supports numeric point sizes"),
    run.font === undefined ? "" : control("f", fontIndex(context, run, "run")),
    run.color === undefined ? "" : control("cf", colorIndex(context, run, "run")),
    run.characterSpacing === undefined
      ? ""
      : control(
          "expndtw",
          typeof run.characterSpacing === "number"
            ? run.characterSpacing
            : reject(
                "run",
                "run.characterSpacing",
                String(run.characterSpacing),
                "RTF writer supports numeric twip spacing",
              ),
        ),
    run.kern === undefined
      ? ""
      : control(
          "kerning",
          typeof run.kern === "number"
            ? Math.round(run.kern / 10)
            : reject(
                "run",
                "run.kern",
                String(run.kern),
                "RTF writer supports numeric twip kerning",
              ),
        ),
    run.allCaps === true ? control("caps") : "",
    run.allCaps === false ? control("caps", 0) : "",
    run.smallCaps === true ? control("scaps") : "",
    run.smallCaps === false ? control("scaps", 0) : "",
    run.outline === true ? control("outl") : "",
    run.outline === false ? control("outl", 0) : "",
    run.shadow === true ? control("shad") : "",
    run.shadow === false ? control("shad", 0) : "",
    run.vanish === true ? control("v") : "",
    run.vanish === false ? control("v", 0) : "",
    run.emboss === true ? control("impr") : "",
    run.emboss === false ? control("impr", 0) : "",
    run.scale === undefined ? "" : control("charscalex", run.scale),
    (() => {
      if (run.style === undefined) return "";
      const code = styleCode(context, run.style);
      if (code === undefined) {
        reject(
          "run",
          "run.style",
          run.style,
          "RTF character style must be declared in the stylesheet",
        );
      }
      return control("cs", code);
    })(),
    (() => {
      const tag = run.language?.value;
      if (tag === undefined) return "";
      const lcid = tagToLcid(tag);
      if (lcid === undefined) {
        reject("run", "run.language.value", tag, "RTF language table has no such tag");
      }
      return control("lang", lcid);
    })(),
    run.verticalAlign === "superscript"
      ? control("super")
      : run.verticalAlign === "subscript"
        ? control("sub")
        : run.verticalAlign === "baseline"
          ? control("nosupersub")
          : "",
  ].join("");
}

export function writeInlineChild(
  child: NonNullable<RunOptions["children"]>[number],
  context: RtfGenerateContext,
  path: string,
): string {
  if (typeof child === "string") return rtfText(child);
  if ("text" in child && typeof child.text === "string") return rtfText(child.text);
  if ("tab" in child && child.tab === true) return control("tab");
  if ("break" in child) {
    if (child.break !== 1) reject("run", path, "break", "RTF can project only one line/page break");
    return control("line");
  }
  if ("hyperlink" in child) {
    const hyperlink = child.hyperlink;
    assertAllowed(hyperlink, ["url", "children"], "hyperlink", `${path}.hyperlink`);
    if (
      hyperlink.url === undefined ||
      hyperlink.children === undefined ||
      hyperlink.children.length !== 1 ||
      typeof hyperlink.children[0] !== "string"
    ) {
      reject("hyperlink", `${path}.hyperlink`, "children", "RTF fields project one text result");
    }
    const url = hyperlink.url.replace(/"/g, '\\"');
    return `{\\field{\\*\\fldinst HYPERLINK "${url}"}{\\fldrslt ${rtfText(hyperlink.children[0])}}}`;
  }
  if ("bookmarkStart" in child) {
    assertAllowed(child.bookmarkStart, ["id", "name"], "bookmark", `${path}.bookmarkStart`);
    return `{\\*\\bkmkstart ${rtfText(child.bookmarkStart.name)}}`;
  }
  if ("bookmarkEnd" in child) {
    assertAllowed(child.bookmarkEnd, ["id"], "bookmark", `${path}.bookmarkEnd`);
    return "{\\*\\bkmkend }";
  }
  if ("endnoteReference" in child) {
    return writeEndnoteReference(context, child.endnoteReference, `${path}.endnoteReference`);
  }
  const field = writeFieldChild(child, path);
  if (field !== undefined) return field;
  const drawing = writeDrawingChild(child, path);
  if (drawing !== undefined) return drawing;
  reject("run", path, Object.keys(child)[0] ?? "child", "RTF cannot represent this inline child");
}

export function writeRun(run: RunOptions, context: RtfGenerateContext, path: string): string {
  const properties = writeRunProperties(run, context);
  if (run.text !== undefined) {
    assertAllowed(run, [...RUN_PROPERTIES, "text"], "run", path);
    if (!properties) return rtfText(run.text);
    return `{${properties}${rtfText(run.text)}}`;
  }
  const children = (run.children ?? []).map((child, index) =>
    writeInlineChild(child, context, `${path}.children[${index}]`),
  );
  if (!properties) return children.join("");
  return `{${properties}${children.join("")}}`;
}
