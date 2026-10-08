import type { RunOptions } from "@office-open/docx";

import { tagToLcid } from "../langs";
import { assertAllowed, colorIndex, fontIndex, reject, type RtfGenerateContext } from "./context";
import { control, rtfText } from "./escape";
import { writeFieldChild } from "./field";
import { writeDrawingChild } from "./shape-picture";

const RUN_PROPERTIES = [
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
] as const;

export function writeRunProperties(run: RunOptions, context: RtfGenerateContext): string {
  assertAllowed(run, [...RUN_PROPERTIES, "text", "children"], "run", "run");
  return [
    run.bold === true ? control("b") : "",
    run.bold === false ? control("b", 0) : "",
    run.italic === true ? control("i") : "",
    run.italic === false ? control("i", 0) : "",
    run.strike === true ? control("strike") : "",
    run.strike === false ? control("strike", 0) : "",
    run.underline
      ? (assertAllowed(run.underline, ["type"], "run", "run.underline"), control("ul"))
      : "",
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
    run.style === undefined
      ? ""
      : run.style.match(/^rtf-character-style-(\d+)$/)
        ? control("cs", Number(run.style.match(/^rtf-character-style-(\d+)$/)![1]))
        : reject("run", "run.style", run.style, "RTF writer only projects parsed character styles"),
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
    writeInlineChild(child, `${path}.children[${index}]`),
  );
  if (!properties) return children.join("");
  return `{${properties}${children.join("")}}`;
}
