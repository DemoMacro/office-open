import type { RunOptions, StylesOptions } from "@office-open/docx";

import {
  assertAllowed,
  registerColor,
  reject,
  reserveStyleCode,
  styleCode,
  type RtfGenerateContext,
} from "./context";
import { control, rtfText } from "./escape";
import { writeRunProperties } from "./run";

type StyleOptions = Omit<
  NonNullable<StylesOptions["paragraphStyles"]>[number] &
    NonNullable<StylesOptions["characterStyles"]>[number],
  "type"
>;
type Entry = { style: StyleOptions; type: "paragraph" | "character" };

function paragraphProperties(paragraph: object, context: RtfGenerateContext, path: string): string {
  assertAllowed(
    paragraph,
    ["alignment", "indent", "spacing", "tabStops", "border", "shading"],
    "styles",
    `${path}.paragraph`,
  );
  const options = paragraph as {
    alignment?: string;
    indent?: Record<string, number | string | undefined>;
    spacing?: Record<string, number | string | undefined>;
    tabStops?: { type: string; position: number | string }[];
    border?: Record<string, { style?: string; size?: number; color?: string } | false | undefined>;
    shading?: { type?: string; fill?: string };
  };
  let output =
    options.alignment === "left"
      ? control("ql")
      : options.alignment === "right"
        ? control("qr")
        : options.alignment === "center"
          ? control("qc")
          : options.alignment === "both"
            ? control("qj")
            : "";
  for (const [field, name] of [
    ["left", "li"],
    ["right", "ri"],
    ["firstLine", "fi"],
  ] as const) {
    const value = options.indent?.[field];
    if (value === undefined) continue;
    if (typeof value !== "number")
      reject("styles", `${path}.paragraph.indent.${field}`, String(value), "RTF requires twips");
    output += control(name, value);
  }
  for (const [field, name] of [
    ["before", "sb"],
    ["after", "sa"],
    ["line", "sl"],
  ] as const) {
    const value = options.spacing?.[field];
    if (value === undefined) continue;
    if (typeof value !== "number")
      reject("styles", `${path}.paragraph.spacing.${field}`, String(value), "RTF requires twips");
    output += control(name, value);
  }
  for (const stop of options.tabStops ?? []) {
    assertAllowed(stop, ["type", "position"], "styles", `${path}.paragraph.tabStops`);
    if (stop.type !== "left" || typeof stop.position !== "number")
      reject(
        "styles",
        `${path}.paragraph.tabStops`,
        stop.type,
        "RTF supports left numeric tab stops",
      );
    output += control("tx", stop.position);
  }
  for (const [edge, name] of [
    ["top", "brdrt"],
    ["bottom", "brdrb"],
    ["left", "brdrl"],
    ["right", "brdrr"],
  ] as const) {
    const border = options.border?.[edge];
    if (!border) continue;
    assertAllowed(border, ["style", "size", "color"], "styles", `${path}.paragraph.border.${edge}`);
    if (border.style !== "single")
      reject(
        "styles",
        `${path}.paragraph.border.${edge}`,
        String(border.style),
        "RTF supports single borders",
      );
    output += control(name);
    if (border.size !== undefined) output += control("brdrw", border.size);
    if (border.color !== undefined)
      output += control("brdrcf", registerColor(context, border.color));
  }
  if (options.shading) {
    assertAllowed(options.shading, ["type", "fill"], "styles", `${path}.paragraph.shading`);
    if (options.shading.type !== "clear" || typeof options.shading.fill !== "string")
      reject("styles", `${path}.paragraph.shading`, "shading", "RTF supports clear string fills");
    output += control("cbpat", registerColor(context, options.shading.fill));
  }
  return output;
}

function runProperties(run: object, context: RtfGenerateContext, path: string): string {
  void path;
  const style = "style" in run ? run.style : undefined;
  if (typeof style === "string" && styleCode(context, style) === undefined) {
    reject("styles", "run.style", style, "RTF character style must be declared");
  }
  return writeRunProperties(run as RunOptions, context);
}

function referencedCode(context: RtfGenerateContext, id: string, path: string): number {
  const code = styleCode(context, id);
  if (code === undefined) reject("styles", path, id, "RTF style reference must be declared");
  return code;
}

export function writeStyles(
  styles: StylesOptions | undefined,
  context: RtfGenerateContext,
): string {
  if (!styles) return "";
  assertAllowed(styles, ["paragraphStyles", "characterStyles"], "styles", "styles");
  const entries: Entry[] = [
    ...(styles.paragraphStyles ?? []).map((style) => ({ style, type: "paragraph" as const })),
    ...(styles.characterStyles ?? []).map((style) => ({ style, type: "character" as const })),
  ];
  for (const { style } of entries) {
    if (style.id === undefined) reject("styles", "id", "id", "RTF requires a style id");
    reserveStyleCode(context, style.id);
  }
  const output = entries
    .map(({ style, type }, index) => {
      const path = type === "paragraph" ? `paragraphStyles[${index}]` : `characterStyles[${index}]`;
      assertAllowed(
        style,
        [
          "id",
          "name",
          "basedOn",
          "next",
          "link",
          "hidden",
          "uiPriority",
          "quickFormat",
          "paragraph",
          "run",
        ],
        "styles",
        path,
      );
      if (style.id === undefined) reject("styles", path, "id", "RTF requires a style id");
      const code = styleCode(context, style.id);
      if (code === undefined) reject("styles", path, "id", "RTF requires a style id");
      let body = control(type === "character" ? "cs" : "s", code);
      body +=
        type === "paragraph"
          ? paragraphProperties(style.paragraph ?? {}, context, path)
          : runProperties(style.run ?? {}, context, path);
      body += [
        style.basedOn === undefined
          ? ""
          : control("sbasedon", referencedCode(context, style.basedOn, `${path}.basedOn`)),
        style.next === undefined
          ? ""
          : control("snext", referencedCode(context, style.next, `${path}.next`)),
        style.link === undefined
          ? ""
          : control("slink", referencedCode(context, style.link, `${path}.link`)),
        style.hidden ? control("shidden") : "",
        style.uiPriority === undefined ? "" : control("spriority", style.uiPriority),
        style.quickFormat ? control("sqformat") : "",
      ].join("");
      return `${body} ${rtfText((style.name ?? style.id).replace(/;$/u, ""))};`;
    })
    .join("");
  return output ? `{\\stylesheet${output}}` : "";
}
