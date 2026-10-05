import type { StylesOptions } from "@office-open/docx";

import { assertAllowed, registerColor, reject, type RtfGenerateContext } from "./context";
import { control, rtfText } from "./escape";

type StyleOptions = Omit<
  NonNullable<StylesOptions["paragraphStyles"]>[number] &
    NonNullable<StylesOptions["characterStyles"]>[number],
  "type"
>;
type Entry = { style: StyleOptions; type: "paragraph" | "character" };

function code(id: string, codes: Map<string, number>): number {
  const known = codes.get(id);
  if (known !== undefined) return known;
  const parsed = /(?:^rtf-style-|^rtf-character-style-)(\d+)$/u.exec(id)?.[1];
  if (parsed === undefined) reject("styles", "id", id, "RTF requires a canonical rtf style id");
  return Number(parsed);
}

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
  assertAllowed(
    run,
    ["bold", "italic", "strike", "verticalAlign", "underline", "size", "color", "font"],
    "styles",
    `${path}.run`,
  );
  const options = run as {
    bold?: boolean;
    italic?: boolean;
    strike?: boolean;
    verticalAlign?: "baseline" | "subscript" | "superscript";
    underline?: object;
    size?: number | string;
    color?: string | object;
    font?: string | object;
  };
  let output = "";
  if (options.bold !== undefined) output += control("b", options.bold ? 1 : 0);
  if (options.italic !== undefined) output += control("i", options.italic ? 1 : 0);
  if (options.underline) {
    assertAllowed(options.underline, ["type"], "styles", `${path}.run.underline`);
    output += control("ul");
  }
  if (options.size !== undefined) {
    if (typeof options.size !== "number")
      reject("styles", `${path}.run.size`, String(options.size), "RTF requires points");
    output += control("fs", Math.round(options.size * 2));
  }
  if (options.color !== undefined) {
    if (typeof options.color !== "string")
      reject("styles", `${path}.run.color`, "value", "RTF requires a hex color");
    output += control("cf", registerColor(context, options.color));
  }
  if (options.font !== undefined) {
    if (typeof options.font !== "string")
      reject("styles", `${path}.run.font`, "value", "RTF requires a font name");
    const index = context.fonts.indexOf(options.font);
    output += control("f", index === -1 ? context.fonts.push(options.font) - 1 : index);
  }
  return output;
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
  let next = 0;
  const codes = new Map<string, number>();
  for (const { style } of entries) {
    const parsed = /(?:^rtf-style-|^rtf-character-style-)(\d+)$/u.exec(style.id ?? "")?.[1];
    const value = parsed === undefined ? next++ : Number(parsed);
    if (style.id !== undefined) codes.set(style.id, value);
    next = Math.max(next + 1, value + 1);
  }
  const output = entries
    .map(({ style, type }, index) => {
      const path = type === "paragraph" ? `paragraphStyles[${index}]` : `characterStyles[${index}]`;
      assertAllowed(
        style,
        ["id", "name", "basedOn", "next", "link", "hidden", "uiPriority", "quickFormat", type],
        "styles",
        path,
      );
      if (style.id === undefined) reject("styles", path, "id", "RTF requires a style id");
      let body = control(type === "character" ? "cs" : "s", code(style.id, codes));
      body +=
        type === "paragraph"
          ? paragraphProperties(style.paragraph ?? {}, context, path)
          : runProperties(style.run ?? {}, context, path);
      body += [
        style.basedOn === undefined ? "" : control("sbasedon", code(style.basedOn, codes)),
        style.next === undefined ? "" : control("snext", code(style.next, codes)),
        style.link === undefined ? "" : control("slink", code(style.link, codes)),
        style.hidden ? control("shidden") : "",
        style.uiPriority === undefined ? "" : control("spriority", style.uiPriority),
        style.quickFormat ? control("sqformat") : "",
      ].join("");
      return `${body} ${rtfText((style.name ?? style.id).replace(/;$/u, ""))};`;
    })
    .join("");
  return output ? `{\\stylesheet${output}}` : "";
}
