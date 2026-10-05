import type {
  ParagraphStylePropertiesOptions,
  RunStylePropertiesOptions,
  StylesOptions,
} from "@office-open/docx";

import type { RtfToken } from "../tokenizer";

export type StyleDraft = {
  code: number;
  type: "paragraph" | "character";
  name: string;
  basedOn?: number;
  link?: number;
  next?: number;
  hidden?: boolean;
  priority?: number;
  quickFormat?: boolean;
  paragraph?: ParagraphStylePropertiesOptions;
  run?: RunStylePropertiesOptions;
};

const decode = (tokens: readonly RtfToken[], start: number, end: number): string => {
  let value = "";
  for (let index = start; index < end; index += 1) {
    const token = tokens[index];
    if (token?.kind === "text") value += token.value;
    else if (token?.kind === "hex") value += String.fromCharCode(Number.parseInt(token.value, 16));
    else if (token?.kind === "control" && token.word === "u" && token.param !== undefined)
      value += String.fromCodePoint(token.param < 0 ? token.param + 65536 : token.param);
  }
  return value;
};

const applyParagraph = (
  paragraph: ParagraphStylePropertiesOptions,
  word: string,
  parameter: number | undefined,
  colors: readonly (string | null)[],
): boolean => {
  if (word === "ql") paragraph.alignment = "left";
  else if (word === "qr") paragraph.alignment = "right";
  else if (word === "qc") paragraph.alignment = "center";
  else if (word === "qj") paragraph.alignment = "both";
  else if (word === "li" && parameter !== undefined)
    paragraph.indent = { ...paragraph.indent, left: parameter };
  else if (word === "ri" && parameter !== undefined)
    paragraph.indent = { ...paragraph.indent, right: parameter };
  else if (word === "fi" && parameter !== undefined)
    paragraph.indent = { ...paragraph.indent, firstLine: parameter };
  else if (word === "sb" && parameter !== undefined)
    paragraph.spacing = { ...paragraph.spacing, before: parameter };
  else if (word === "sa" && parameter !== undefined)
    paragraph.spacing = { ...paragraph.spacing, after: parameter };
  else if (word === "sl" && parameter !== undefined)
    paragraph.spacing = { ...paragraph.spacing, line: parameter };
  else if (word === "tx" && parameter !== undefined)
    paragraph.tabStops = [...(paragraph.tabStops ?? []), { type: "left", position: parameter }];
  else if (word === "cbpat" && parameter !== undefined)
    paragraph.shading = { type: "clear", fill: colors[parameter] ?? "auto" };
  else if (word === "brdrt" || word === "brdrb" || word === "brdrl" || word === "brdrr") {
    paragraph.border ??= {};
    paragraph.border[
      word === "brdrt" ? "top" : word === "brdrb" ? "bottom" : word === "brdrl" ? "left" : "right"
    ] = { style: "single", ...(parameter !== undefined ? { size: parameter } : {}) };
  } else return false;
  return true;
};

const applyRun = (
  run: RunStylePropertiesOptions,
  word: string,
  parameter: number | undefined,
  colors: readonly (string | null)[],
  fonts: ReadonlyMap<number, string>,
): boolean => {
  if (word === "b") run.bold = parameter !== 0;
  else if (word === "i") run.italic = parameter !== 0;
  else if (word === "ul" || word === "ulw") run.underline = { type: "single" };
  else if (word === "ulnone") run.underline = undefined;
  else if (word === "fs" && parameter !== undefined && parameter > 0) run.size = parameter / 2;
  else if (word === "cf" && parameter !== undefined) run.color = colors[parameter] ?? undefined;
  else if (word === "f" && parameter !== undefined) run.font = fonts.get(parameter);
  else return false;
  return true;
};

export function parseStylesheet(
  tokens: readonly RtfToken[],
  start: number,
  end: number,
  colors: readonly (string | null)[],
  fonts: ReadonlyMap<number, string>,
): StyleDraft[] {
  const styles: StyleDraft[] = [];
  let current: StyleDraft | undefined;
  let depth = 0;
  let nameStart: number | undefined;
  let borderEdge: "top" | "bottom" | "left" | "right" | undefined;
  const finish = () => {
    if (!current) return;
    if (nameStart !== undefined)
      current.name = decode(tokens, nameStart, end).replace(/;.*$/u, "").trim();
    if (current.name) styles.push(current);
    current = undefined;
    nameStart = undefined;
  };
  for (let index = start; index < end; index += 1) {
    const token = tokens[index];
    if (!token) continue;
    if (token.kind === "group-start") {
      depth += 1;
      continue;
    }
    if (token.kind === "group-end") {
      if (depth === 0) break;
      depth -= 1;
      if (depth === 0) finish();
      continue;
    }
    if (token.kind !== "control") {
      if (current && token.kind === "text" && nameStart === undefined) nameStart = index;
      continue;
    }
    const { word, param } = token;
    if (word === undefined) continue;
    if (word === "s" || word === "cs") {
      finish();
      if (param !== undefined)
        current = { code: param, type: word === "cs" ? "character" : "paragraph", name: "" };
      continue;
    }
    if (!current) continue;
    if (word === "sbasedon" && param !== undefined) current.basedOn = param;
    else if (word === "slink" && param !== undefined) current.link = param;
    else if (word === "snext" && param !== undefined) current.next = param;
    else if (word === "shidden") current.hidden = true;
    else if (word === "spriority" && param !== undefined) current.priority = param;
    else if (word === "sqformat") current.quickFormat = true;
    else if (
      current.type === "paragraph" &&
      (word === "brdrw" || word === "brdrcf") &&
      borderEdge
    ) {
      const border = current.paragraph!.border![borderEdge]!;
      if (word === "brdrw" && param !== undefined) border.size = param;
      else if (word === "brdrcf" && param !== undefined) border.color = colors[param] ?? undefined;
    } else if (
      current.type === "paragraph" &&
      applyParagraph((current.paragraph ??= {}), word, param, colors)
    ) {
      borderEdge =
        word === "brdrt"
          ? "top"
          : word === "brdrb"
            ? "bottom"
            : word === "brdrl"
              ? "left"
              : word === "brdrr"
                ? "right"
                : undefined;
    } else {
      borderEdge = undefined;
      applyRun((current.run ??= {}), word, param, colors, fonts);
    }
  }
  finish();
  return styles;
}

export const styleId = (draft: StyleDraft): string =>
  draft.type === "character" ? `rtf-character-style-${draft.code}` : `rtf-style-${draft.code}`;

export function stylesFromDrafts(drafts: readonly StyleDraft[]): StylesOptions | undefined {
  if (drafts.length === 0) return undefined;
  const ids = new Map(drafts.map((draft) => [draft.code, styleId(draft)]));
  const map = (draft: StyleDraft) => ({
    id: styleId(draft),
    ...(draft.name ? { name: draft.name } : {}),
    ...(draft.basedOn !== undefined && ids.has(draft.basedOn)
      ? { basedOn: ids.get(draft.basedOn) }
      : {}),
    ...(draft.next !== undefined && ids.has(draft.next) ? { next: ids.get(draft.next) } : {}),
    ...(draft.link !== undefined && ids.has(draft.link) ? { link: ids.get(draft.link) } : {}),
    ...(draft.hidden ? { hidden: true } : {}),
    ...(draft.priority !== undefined ? { uiPriority: draft.priority } : {}),
    ...(draft.quickFormat ? { quickFormat: true } : {}),
    ...(draft.type === "paragraph" && draft.paragraph ? { paragraph: draft.paragraph } : {}),
    ...(draft.run ? { run: draft.run } : {}),
  });
  const paragraphStyles = drafts.filter((draft) => draft.type === "paragraph").map(map);
  const characterStyles = drafts.filter((draft) => draft.type === "character").map(map);
  return {
    ...(paragraphStyles.length ? { paragraphStyles } : {}),
    ...(characterStyles.length ? { characterStyles } : {}),
  };
}
