import type { EndnoteOptions, FootnoteOptions } from "@office-open/docx";

import { assertAllowed, reject, type RtfGenerateContext } from "./context";
import { control, rtfText } from "./escape";

function noteText(note: FootnoteOptions | EndnoteOptions, path: string): string {
  assertAllowed(note, ["id", "children", "autoReferenceMark"], "footnote", path);
  if ("autoReferenceMark" in note && note.autoReferenceMark === false) {
    reject("footnote", path, "autoReferenceMark", "RTF always projects a note reference mark");
  }
  return note.children
    .map((child, childIndex) => {
      const childPath = `${path}.children[${childIndex}]`;
      if (
        typeof child !== "object" ||
        child === null ||
        !("paragraph" in child) ||
        typeof child.paragraph === "string"
      ) {
        reject("footnote", childPath, "child", "RTF footnotes contain paragraphs");
      }
      assertAllowed(child.paragraph, ["children"], "footnote", `${childPath}.paragraph`);
      return (child.paragraph.children ?? [])
        .map((run, runIndex) => {
          assertAllowed(run, ["text"], "footnote", `${childPath}.paragraph.children[${runIndex}]`);
          return rtfText((run as { text?: string }).text ?? "");
        })
        .join("");
    })
    .map((text) => `${text}${control("par")}`)
    .join("");
}

export function writeFootnoteReference(
  context: RtfGenerateContext,
  reference: number | unknown,
  path: string,
): string {
  if (typeof reference !== "number" || !Number.isInteger(reference) || reference < 0) {
    reject("run", path, "footnoteReference", "RTF requires an ordered footnote index");
  }
  const footnote = context.footnotes?.[reference];
  if (!footnote) {
    reject("run", path, "footnoteReference", "RTF footnote reference has no matching content");
  }
  return `{\\footnote ${noteText(footnote, `footnotes[${reference}]`)}}`;
}

export function writeEndnoteReference(
  context: RtfGenerateContext,
  reference: unknown,
  path: string,
): string {
  const index =
    typeof reference === "number"
      ? reference
      : typeof reference === "object" &&
          reference !== null &&
          "id" in reference &&
          typeof reference.id === "number"
        ? reference.id
        : undefined;
  if (typeof index !== "number" || !Number.isInteger(index) || index < 0) {
    reject("run", path, "endnoteReference", "RTF requires an ordered endnote index");
  }
  const endnote = context.endnotes?.[index];
  if (endnote === undefined) {
    reject("run", path, "endnoteReference", "RTF endnote reference has no matching content");
  }
  return `{\\footnote\\ftnalt ${noteText(endnote, `endnotes[${index}]`)}}`;
}
