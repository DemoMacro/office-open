/**
 * TextRun descriptor (a:r / CT_RegularTextRun).
 *
 * @module
 */

import { findChild, escapeXml } from "@office-open/xml";

import type { CustomDescriptor } from "../../descriptor";
import { runPropertiesDesc } from "./run-properties";
import type { Mutable } from "./run-properties";
import type { TextRunOptions } from "./types";

/**
 * Stringify a text-only run — `<a:r><a:t>…</a:t></a:r>` with no `a:rPr`.
 * Shared with the paragraph descriptor, whose string-children shorthand
 * expands to exactly this and would otherwise pay a full run-properties
 * scan plus a throwaway `{ text }` object per child.
 */
export function stringifyTextRun(text: string): string {
  return `<a:r><a:t>${escapeXml(text)}</a:t></a:r>`;
}

export const textRunDesc: CustomDescriptor<TextRunOptions> = {
  kind: "custom",

  stringify(opts, ctx) {
    const body = runPropertiesDesc.stringify(opts, ctx) ?? "";
    // Empty string keeps an explicit <a:t/> — sources carry empty text runs
    // (a:br neighbors) and round-trip must not drop the element.
    if (opts.text !== undefined) {
      return body
        ? `<a:r>${body}<a:t>${escapeXml(opts.text)}</a:t></a:r>`
        : stringifyTextRun(opts.text);
    }
    return body ? `<a:r>${body}</a:r>` : "<a:r/>";
  },

  parse(el, _ctx) {
    const result: Mutable<TextRunOptions> = {};

    const rPr = findChild(el, "a:rPr");
    if (rPr) {
      Object.assign(result, runPropertiesDesc.parse(rPr, _ctx));
    }

    const t = findChild(el, "a:t");
    if (t) {
      result.text = (t.elements ?? [])
        .filter((e) => e.type === "text")
        .map((e) => e.text ?? "")
        .join("");
    }

    return result as TextRunOptions;
  },
};
