import { convertToTwip } from "@office-open/core";
import type { ParagraphOptions, RunOptions } from "@office-open/docx";

import {
  assertAllowed,
  registerColor,
  reject,
  styleCode,
  type RtfGenerateContext,
} from "./context";
import { control, rtfText } from "./escape";
import { writeEndnoteReference, writeFootnoteReference } from "./footnote";
import { writeParagraphNumbering, type NumberingReferences } from "./numbering";
import { writeInlineChild, writeRun } from "./run";

const PARAGRAPH_PROPERTIES = [
  "text",
  "preserveSpace",
  "children",
  "style",
  "alignment",
  "shading",
  "border",
  "numbering",
  "indent",
] as const;

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

function indentControl(
  value: number | NonNullable<NonNullable<ParagraphOptions["indent"]>["left"]>,
  word: string,
  path: string,
): string {
  const twips = convertToTwip(value);
  if (!Number.isInteger(twips)) {
    reject("paragraph", path, word, "RTF indentation requires whole twips");
  }
  return control(word, twips);
}

function writeIndent(indent: NonNullable<ParagraphOptions["indent"]>, path: string): string {
  assertAllowed(
    indent,
    ["left", "start", "right", "end", "firstLine", "hanging"],
    "paragraph",
    path,
  );
  let output = "";
  const left = indent.left ?? indent.start;
  const right = indent.right ?? indent.end;
  if (left !== undefined) output += indentControl(left, "li", path);
  if (right !== undefined) output += indentControl(right, "ri", path);
  if (indent.firstLine !== undefined) output += indentControl(indent.firstLine, "fi", path);
  if (indent.hanging !== undefined) {
    output += indentControl(-convertToTwip(indent.hanging), "fi", path);
  }
  return output;
}

export function writeParagraph(
  paragraph: ParagraphOptions,
  context: RtfGenerateContext,
  references: NumberingReferences,
  path: string,
): string {
  assertAllowed(paragraph, PARAGRAPH_PROPERTIES, "paragraph", path);
  let properties = "";
  if (paragraph.style !== undefined) {
    const styleCodeValue = styleCode(context, paragraph.style);
    if (styleCodeValue === undefined) {
      reject("paragraph", "style", paragraph.style, "RTF paragraph style must be declared");
    }
    properties += control("s", styleCodeValue);
  }
  if (paragraph.alignment) properties += alignmentControl(paragraph.alignment);
  if (paragraph.numbering) {
    properties += writeParagraphNumbering(paragraph.numbering, references, path);
  }
  if (paragraph.indent) {
    properties += writeIndent(paragraph.indent, `${path}.indent`);
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
  if (paragraph.text !== undefined && paragraph.children !== undefined) {
    reject("paragraph", "children", "children", "text and children are mutually exclusive");
  }
  const childrenText =
    paragraph.text !== undefined
      ? writeRun({ text: paragraph.text }, context, `${path}.text`)
      : (paragraph.children ?? [])
          .map((child, index) => {
            const childPath = `${path}.children[${index}]`;
            if (typeof child === "string") return rtfText(child);
            if ("footnoteReference" in child) {
              return writeFootnoteReference(context, child.footnoteReference, childPath);
            }
            if ("endnoteReference" in child) {
              return writeEndnoteReference(context, child.endnoteReference, childPath);
            }
            const isRun =
              "text" in child ||
              "bold" in child ||
              "italic" in child ||
              "underline" in child ||
              "strike" in child ||
              "verticalAlign" in child ||
              "color" in child ||
              "size" in child ||
              "font" in child ||
              "children" in child;
            return isRun
              ? writeRun(child as RunOptions, context, childPath)
              : writeInlineChild(child, context, childPath);
          })
          .join("");
  return `${properties}${childrenText}${control("par")}`;
}
