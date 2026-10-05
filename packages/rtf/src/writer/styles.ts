import type { StylesOptions } from "@office-open/docx";

import { assertAllowed } from "./context";
import { control, rtfText } from "./escape";

type StyleDefinition = { id: string; name: string; type: "paragraph" | "character" };

export function writeStyles(styles: StylesOptions | undefined): string {
  if (!styles) return "";
  assertAllowed(styles, ["paragraphStyles", "characterStyles"], "styles", "styles");
  const definitions: StyleDefinition[] = [];
  for (const style of styles.paragraphStyles ?? []) {
    assertAllowed(style, ["id", "name", "basedOn"], "styles", "paragraphStyles");
    if (typeof style.name !== "string") {
      assertAllowed(style, ["id"], "styles", "paragraphStyles");
    }
    definitions.push({ id: style.id, name: style.name ?? "", type: "paragraph" });
  }
  for (const style of styles.characterStyles ?? []) {
    assertAllowed(style, ["id", "name", "basedOn"], "styles", "characterStyles");
    if (typeof style.name !== "string") {
      assertAllowed(style, ["id"], "styles", "characterStyles");
    }
    definitions.push({ id: style.id, name: style.name ?? "", type: "character" });
  }
  if (definitions.length === 0) return "";
  const entries = definitions
    .map(
      (definition) =>
        `${control(definition.type === "character" ? "cs" : "s", 0)} ${rtfText(definition.name.replace(/;$/u, ""))};`,
    )
    .join("");
  return `{\\stylesheet${entries}}`;
}
