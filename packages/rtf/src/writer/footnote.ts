import type { FootnoteOptions } from "@office-open/docx";

import { assertAllowed, reject, type RtfGenerateContext } from "./context";
import { control, rtfText } from "./escape";

function footnoteText(footnote: FootnoteOptions, path: string): string {
  assertAllowed(footnote, ["children"], "footnote", path);
  return footnote.children
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
  return `{\\footnote ${footnoteText(footnote, `footnotes[${reference}]`)}}`;
}
