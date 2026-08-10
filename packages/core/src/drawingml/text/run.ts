/**
 * TextRun descriptor (a:r / CT_RegularTextRun).
 *
 * @module
 */

import { findChild, escapeXml } from "@office-open/xml";

import type { CustomDescriptor } from "../../descriptor";
import { runPropertiesDesc } from "./run-properties";
import type { Mutable } from "./run-properties";
import type { RunOptions } from "./types";

export const textRunDesc: CustomDescriptor<RunOptions> = {
  kind: "custom",

  stringify(opts, ctx) {
    const body = runPropertiesDesc.stringify(opts, ctx) ?? "";
    if (opts.text) {
      return `<a:r>${body}<a:t>${escapeXml(opts.text)}</a:t></a:r>`;
    }
    return body ? `<a:r>${body}</a:r>` : "<a:r/>";
  },

  parse(el, _ctx) {
    const result: Mutable<RunOptions> = {};

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

    return result as RunOptions;
  },
};
