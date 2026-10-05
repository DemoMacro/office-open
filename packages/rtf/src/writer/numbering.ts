import type { NumberingOptions } from "@office-open/docx";

import { assertAllowed, reject } from "./context";
import { control } from "./escape";

export type NumberingReferenceInfo = { id: number; format: "bullet" | "decimal" };
export type NumberingReferences = Map<string, NumberingReferenceInfo>;

export function writeNumberingTable(
  numbering: NumberingOptions | undefined,
  references: NumberingReferences,
): string {
  if (!numbering) return "";
  assertAllowed(numbering, ["abstractNumberings"], "numbering", "numbering");
  const entries = numbering.abstractNumberings.map((definition, index) => {
    const id = index + 1;
    const format = definition.levels.some((level) => level.format === "bullet")
      ? "bullet"
      : "decimal";
    references.set(definition.reference, { id, format });
    assertAllowed(definition, ["reference", "levels"], "numbering", `abstractNumberings[${index}]`);
    const levels = definition.levels
      .map((level, levelIndex) => {
        assertAllowed(
          level,
          ["level", "format", "text"],
          "numbering",
          `abstractNumberings[${index}].levels[${levelIndex}]`,
        );
        if (level.format !== "decimal" && level.format !== "bullet") {
          reject(
            "numbering",
            `abstractNumberings[${index}].levels[${levelIndex}].format`,
            String(level.format),
            "RTF writer supports decimal and bullet levels",
          );
        }
        return `${control("listlevel", level.level)}${control(level.format === "bullet" ? "pnbullet" : "pndec")}`;
      })
      .join("");
    return `{\\listoverride${control("ls", id)}${levels}}`;
  });
  const overrides = entries.filter((_, index) =>
    numbering.abstractNumberings[index]?.reference.startsWith("rtf-list-override-"),
  );
  return overrides.length > 0 ? `{\\listoverridetable${overrides.join("")}}` : "";
}

export function writeParagraphNumbering(
  numbering: unknown,
  references: NumberingReferences,
  path: string,
): string {
  if (typeof numbering !== "object" || numbering === null || !("reference" in numbering)) {
    reject("paragraph", `${path}.numbering`, "instance", "RTF writer requires a reference");
  }
  const options = numbering as { reference?: string; level?: number };
  assertAllowed(numbering, ["reference", "level"], "paragraph", `${path}.numbering`);
  const reference = options.reference === undefined ? undefined : references.get(options.reference);
  if (reference === undefined) {
    reject(
      "paragraph",
      `${path}.numbering`,
      String(options.reference),
      "numbering reference is not declared",
    );
  }
  return [
    control("ls", reference.id),
    options.level === undefined ? "" : control("ilvl", options.level),
    `{\\pn\\${reference.format === "bullet" ? "pnbullet" : "pndec"}}`,
  ].join("");
}
