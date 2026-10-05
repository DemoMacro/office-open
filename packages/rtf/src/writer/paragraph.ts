import type { ParagraphOptions, RunOptions } from "@office-open/docx";

import { assertAllowed, registerColor, reject, type RtfGenerateContext } from "./context";
import { control, rtfText } from "./escape";
import { writeParagraphNumbering, type NumberingReferences } from "./numbering";
import { writeInlineChild, writeRun } from "./run";

const PARAGRAPH_PROPERTIES = ["children", "alignment", "shading", "border", "numbering"] as const;

function alignmentControl(alignment: NonNullable<ParagraphOptions["alignment"]>): string {
  if (alignment === "left") return control("ql");
  if (alignment === "right") return control("qr");
  if (alignment === "center") return control("qc");
  if (alignment === "both") return control("qj");
  reject(
    "paragraph",
    "alignment",
    alignment,
    "RTF writer supports left, right, center, and justified text",
  );
}

export function writeParagraph(
  paragraph: ParagraphOptions,
  context: RtfGenerateContext,
  references: NumberingReferences,
  path: string,
): string {
  assertAllowed(paragraph, PARAGRAPH_PROPERTIES, "paragraph", path);
  let properties = "";
  if (paragraph.alignment) properties += alignmentControl(paragraph.alignment);
  if (paragraph.numbering) {
    properties += writeParagraphNumbering(paragraph.numbering, references, path);
  }
  if (paragraph.shading) {
    assertAllowed(paragraph.shading, ["type", "fill"], "paragraph", `${path}.shading`);
    if (
      paragraph.shading.type !== "clear" ||
      paragraph.shading.fill === undefined ||
      typeof paragraph.shading.fill !== "string"
    ) {
      reject(
        "paragraph",
        `${path}.shading`,
        paragraph.shading.type ?? "fill",
        "RTF writer supports clear shading with a fill",
      );
    }
    properties +=
      paragraph.shading.fill === "auto"
        ? control("shading")
        : control("cbpat", registerColor(context, paragraph.shading.fill));
  }
  if (paragraph.border) {
    assertAllowed(
      paragraph.border,
      ["top", "bottom", "left", "right"],
      "paragraph",
      `${path}.border`,
    );
    const edges = paragraph.border;
    const controls = [
      ["top", "brdrt"],
      ["bottom", "brdrb"],
      ["left", "brdrl"],
      ["right", "brdrr"],
    ] as const;
    for (const [edge, name] of controls) {
      const value = edges[edge];
      if (!value) continue;
      assertAllowed(value, ["style", "size"], "paragraph", `${path}.border.${edge}`);
      if (value.style !== "single") {
        reject(
          "paragraph",
          `${path}.border.${edge}`,
          value.style,
          "RTF writer supports single borders",
        );
      }
      properties += control(name, value.size);
    }
  }
  const childrenText = (paragraph.children ?? [])
    .map((child, index) => {
      const childPath = `${path}.children[${index}]`;
      if (typeof child === "string") return rtfText(child);
      const isRun =
        "text" in child ||
        "bold" in child ||
        "italic" in child ||
        "underline" in child ||
        "color" in child ||
        "size" in child ||
        "font" in child ||
        "children" in child;
      return isRun
        ? writeRun(child as RunOptions, context, childPath)
        : writeInlineChild(child, childPath);
    })
    .join("");
  return `${properties}${childrenText}${control("par")}`;
}
