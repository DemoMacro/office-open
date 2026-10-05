import type { LevelsOptions, NumberingOptions } from "@office-open/docx";

import type { RtfToken } from "../tokenizer";

type NumberingDefinition = NonNullable<NumberingOptions["abstractNumberings"]>[number];
type NumberingLevel = LevelsOptions;
type NumberFormat =
  | "decimal"
  | "lowerLetter"
  | "upperLetter"
  | "lowerRoman"
  | "upperRoman"
  | "bullet"
  | "ordinal"
  | "cardinalText"
  | "ordinalText"
  | "arabicAlpha"
  | "chicago"
  | "ideographDigital";

export const RTF_NUMBER_FORMAT: Record<NumberFormat, { code: number; control: string }> = {
  decimal: { code: 1, control: "pndec" },
  lowerLetter: { code: 4, control: "pnlcltr" },
  upperLetter: { code: 5, control: "pnucltr" },
  lowerRoman: { code: 2, control: "pnlcrm" },
  upperRoman: { code: 3, control: "pnucrm" },
  bullet: { code: 0, control: "pnbullet" },
  ordinal: { code: 6, control: "pnord" },
  cardinalText: { code: 7, control: "pncardt" },
  ordinalText: { code: 8, control: "pnordt" },
  arabicAlpha: { code: 11, control: "pnarabicalpha" },
  chicago: { code: 22, control: "pnchicago" },
  ideographDigital: { code: 24, control: "pnideographdigital" },
};

const formatFromCode = (code?: number): NumberingLevel["format"] =>
  (Object.entries(RTF_NUMBER_FORMAT).find((entry) => entry[1].code === code)?.[0] ??
    "decimal") as NumberingLevel["format"];
const levelAlignment = (code?: number): NumberingLevel["alignment"] =>
  code === 1 ? "center" : code === 2 ? "right" : code === 3 ? "both" : "left";

const decode = (tokens: readonly RtfToken[], start: number, end: number): string => {
  let value = "";
  for (let index = start; index < end; index += 1) {
    const token = tokens[index];
    if (token?.kind === "text") value += token.value;
    else if (token?.kind === "hex") value += String.fromCharCode(Number.parseInt(token.value, 16));
    else if (token?.kind === "control" && token.word === "u" && token.param !== undefined)
      value += String.fromCodePoint(token.param < 0 ? token.param + 65536 : token.param);
  }
  const first = value.codePointAt(0);
  return (first !== undefined && first >= 1 && first <= 31 ? value.slice(1) : value)
    .replaceAll("\u0000", "%1")
    .replace(/;$/u, "");
};

const groupEnd = (tokens: readonly RtfToken[], start: number, end: number): number => {
  let depth = 0;
  for (let index = start; index < end; index += 1) {
    if (tokens[index]?.kind === "group-start") depth += 1;
    else if (tokens[index]?.kind === "group-end") {
      if (depth === 0) return index;
      depth -= 1;
    }
  }
  return end;
};

const levelEnd = (tokens: readonly RtfToken[], start: number, end: number): number => {
  let depth = 0;
  for (let index = start; index < end; index += 1) {
    const token = tokens[index];
    if (token?.kind === "group-start") depth += 1;
    else if (token?.kind === "group-end") {
      if (depth === 0) return index;
      depth -= 1;
    } else if (token?.kind === "control" && (token.word === "listlevel" || token.word === "list"))
      return index;
  }
  return end;
};

function readLevel(
  tokens: readonly RtfToken[],
  start: number,
  end: number,
  initial: number,
): NumberingLevel {
  const level: NumberingLevel = { level: initial, format: "decimal" };
  for (let index = start; index < end; index += 1) {
    const token = tokens[index];
    if (token?.kind !== "control") continue;
    const { word, param } = token;
    if (word === "levelstartat" && param !== undefined) level.start = param;
    else if (word === "levelnfc" && param !== undefined) level.format = formatFromCode(param);
    else if (word === "leveljc" && param !== undefined) level.alignment = levelAlignment(param);
    else if (word === "levelfollow")
      level.suffix = param === 1 ? "space" : param === 2 ? "nothing" : "tab";
    else if (word === "levelindent" && param !== undefined)
      level.paragraph = { ...level.paragraph, indent: { ...level.paragraph?.indent, left: param } };
    else if (word === "levelspace" && param !== undefined)
      level.paragraph = {
        ...level.paragraph,
        spacing: { ...level.paragraph?.spacing, after: param },
      };
    else if (word === "leveltext") {
      const close = groupEnd(tokens, index + 1, end);
      level.text = decode(tokens, index + 1, close);
      index = close;
    } else if (word === "f" && param !== undefined)
      level.run = { ...level.run, font: String(param) };
    else if (word === "cf" && param !== undefined)
      level.run = { ...level.run, color: String(param) };
  }
  return level;
}

function finishDefinition(
  definition: (NumberingDefinition & { listId?: number }) | undefined,
  definitions: NumberingDefinition[],
  listIds: Map<number, string>,
): void {
  if (!definition) return;
  definition.levels.sort((left, right) => left.level - right.level);
  const listId = definition.listId;
  delete definition.listId;
  definitions.push(definition);
  if (listId !== undefined) listIds.set(listId, definition.reference);
}

export function parseListTable(
  tokens: readonly RtfToken[],
  start: number,
  end: number,
): { definitions: NumberingDefinition[]; listIds: Map<number, string> } {
  const definitions: NumberingDefinition[] = [];
  const listIds = new Map<number, string>();
  let definition: (NumberingDefinition & { listId?: number }) | undefined;
  let level: NumberingLevel | undefined;
  const finishLevel = () => {
    if (definition && level) definition.levels.push(level);
    level = undefined;
  };
  for (let index = start; index < end; index += 1) {
    const token = tokens[index];
    if (token?.kind !== "control") continue;
    const { word, param } = token;
    if (word === "list") {
      finishLevel();
      finishDefinition(definition, definitions, listIds);
      definition = { levels: [], reference: `rtf-list-template-${definitions.length + 1}` };
    } else if (definition && word === "listtemplateid" && param !== undefined)
      definition.reference = `rtf-list-template-${param}`;
    else if (definition && word === "listid" && param !== undefined) definition.listId = param;
    else if (definition && word === "listname") {
      const close = groupEnd(tokens, index + 1, end);
      definition.reference = decode(tokens, index + 1, close);
      index = close;
    } else if (word === "listlevel") {
      finishLevel();
      const close = levelEnd(tokens, index + 1, end);
      level = readLevel(tokens, index + 1, close, param ?? definition?.levels.length ?? 0);
      index = close - 1;
      finishLevel();
    }
  }
  finishLevel();
  finishDefinition(definition, definitions, listIds);
  return { definitions, listIds };
}

export function parseListOverrides(
  tokens: readonly RtfToken[],
  start: number,
  end: number,
  listIds: ReadonlyMap<number, string>,
): { definitions: NumberingDefinition[]; listInstances: Map<number, string> } {
  const definitions: NumberingDefinition[] = [];
  const listInstances = new Map<number, string>();
  let listNumber: number | undefined;
  let sourceId: number | undefined;
  const finish = () => {
    if (listNumber === undefined) return;
    const source = sourceId === undefined ? undefined : listIds.get(sourceId);
    const reference = source ?? `rtf-list-${listNumber}`;
    listInstances.set(listNumber, reference);
    if (!source)
      definitions.push({ reference, levels: [{ level: 0, format: "decimal", text: "%1." }] });
  };
  for (let index = start; index < end; index += 1) {
    const token = tokens[index];
    if (token?.kind !== "control") continue;
    if (token.word === "listoverride") {
      finish();
      listNumber = undefined;
      sourceId = undefined;
    } else if (token.word === "ls" && token.param !== undefined) listNumber = token.param;
    else if (token.word === "listid" && token.param !== undefined) sourceId = token.param;
  }
  finish();
  return { definitions, listInstances };
}

export function numberingFromParsed(
  definitions: readonly NumberingDefinition[],
  references: ReadonlySet<string>,
): NumberingOptions | undefined {
  const known = new Set(definitions.map((definition) => definition.reference));
  const fallbacks = [...references]
    .filter((reference) => !known.has(reference))
    .map((reference) => ({
      reference,
      levels: [{ level: 0, format: "decimal" as const, text: "%1." }],
    }));
  return definitions.length || fallbacks.length
    ? { abstractNumberings: [...definitions, ...fallbacks] }
    : undefined;
}
