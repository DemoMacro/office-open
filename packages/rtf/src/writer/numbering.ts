import type { NumberingOptions } from "@office-open/docx";

import { RTF_NUMBER_FORMAT } from "../mappers/numbering";
import { assertAllowed, reject } from "./context";
import { control } from "./escape";

export type NumberingReferenceInfo = { id: number; format: keyof typeof RTF_NUMBER_FORMAT };
export type NumberingReferences = Map<string, NumberingReferenceInfo>;

const entry = (format: string) => {
  const value = RTF_NUMBER_FORMAT[format as keyof typeof RTF_NUMBER_FORMAT];
  if (!value)
    reject("numbering", "format", format, "RTF writer does not support this number format");
  return value;
};

const levelText = (text: string): string => {
  const placeholders = (text.match(/%\d+/gu) ?? []).length;
  const visible = placeholders + text.replace(/%\d+/gu, "").length;
  return `${control("leveltext")}\\'${visible.toString(16).padStart(2, "0")}${text.replace(/%(\d+)/gu, (_, value: string) => `\\'0${Number(value) - 1}`)};`;
};

function level(
  level: NonNullable<NumberingOptions["abstractNumberings"]>[number]["levels"][number],
  path: string,
): string {
  assertAllowed(
    level,
    ["level", "format", "text", "alignment", "start", "paragraph", "suffix", "run"],
    "numbering",
    path,
  );
  if (level.format === undefined) return "";
  const format = entry(level.format);
  const properties = [
    control("levelnfc", format.code),
    control("levelnfcn", format.code),
    level.alignment === undefined
      ? ""
      : control(
          "leveljc",
          level.alignment === "center"
            ? 1
            : level.alignment === "right"
              ? 2
              : level.alignment === "both"
                ? 3
                : 0,
        ),
    level.suffix === undefined
      ? ""
      : control("levelfollow", level.suffix === "space" ? 1 : level.suffix === "nothing" ? 2 : 0),
    level.start === undefined ? "" : control("levelstartat", level.start),
    typeof level.paragraph?.indent?.left === "number"
      ? control("levelindent", level.paragraph.indent.left)
      : "",
  ].join("");
  return `{${control("listlevel", level.level)}{${properties}${levelText(level.text ?? "%1.")}}{${control("levelnumbers")};}}`;
}

export function writeNumberingTable(
  numbering: NumberingOptions | undefined,
  references: NumberingReferences,
): string {
  if (!numbering) return "";
  assertAllowed(numbering, ["abstractNumberings"], "numbering", "numbering");
  if (numbering.abstractNumberings.length === 0) return "";
  const list = numbering.abstractNumberings
    .map((definition, index) => {
      const id = index + 1;
      const format = (definition.levels[0]?.format ?? "decimal") as keyof typeof RTF_NUMBER_FORMAT;
      entry(format);
      references.set(definition.reference, { id, format });
      assertAllowed(
        definition,
        ["reference", "levels"],
        "numbering",
        `abstractNumberings[${index}]`,
      );
      const levels = definition.levels
        .map((value, levelIndex) =>
          level(value, `abstractNumberings[${index}].levels[${levelIndex}]`),
        )
        .join("");
      return `{${control("list")}${control("listtemplateid", id)}${control("listhybrid")}${levels}{${control("listname")}${definition.reference};}${control("listid", id)}}`;
    })
    .join("");
  const overrides = numbering.abstractNumberings
    .map((_, index) => {
      const id = index + 1;
      return `{${control("listoverride")}${control("listid", id)}${control("listoverridecount", 0)}${control("ls", id)}}`;
    })
    .join("");
  return `{\\listtable${list}}{\\listoverridetable${overrides}}`;
}

export function writeParagraphNumbering(
  numbering: unknown,
  references: NumberingReferences,
  path: string,
): string {
  if (typeof numbering !== "object" || numbering === null || !("reference" in numbering))
    reject("paragraph", `${path}.numbering`, "instance", "RTF writer requires a reference");
  const options = numbering as { reference?: string; level?: number };
  assertAllowed(numbering, ["reference", "level"], "paragraph", `${path}.numbering`);
  const reference = options.reference === undefined ? undefined : references.get(options.reference);
  if (reference === undefined)
    reject(
      "paragraph",
      `${path}.numbering`,
      String(options.reference),
      "numbering reference is not declared",
    );
  entry(reference!.format);
  return [
    control("ls", reference!.id),
    options.level === undefined ? "" : control("ilvl", options.level),
    `{\\pntext ${reference!.format === "bullet" ? "•" : "1."}}`,
  ].join("");
}
