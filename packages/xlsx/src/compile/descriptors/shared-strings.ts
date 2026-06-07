/**
 * SharedStrings descriptor for XLSX — generates xl/sharedStrings.xml.
 *
 * The SharedStrings class is an accumulator (register during compilation,
 * serialize at the end). This descriptor handles the serialization/deserialization
 * while the accumulator stays in XlsxWriteContext.
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { findChild, textOf, escapeXml } from "@office-open/xml";

import { buildRstXml } from "../../file/shared-strings";
import type { RichTextOptions } from "../../file/worksheet";

// ── Types ──

/** Serializable snapshot of the shared string table. */
export interface SharedStringsDocOptions {
  /** All entries (plain strings and rich text), in registration order. */
  entries: (string | RichTextOptions)[];
  /** Number of unique plain-string entries (for uniqueCount attribute). */
  uniqueCount: number;
}

// ── Descriptor ──

export const sharedStringsDesc: CustomDescriptor<SharedStringsDocOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    if (opts.entries.length === 0) return undefined;

    const p: string[] = [
      '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"',
      ` count="${opts.entries.length}" uniqueCount="${opts.uniqueCount}">`,
    ];

    for (const entry of opts.entries) {
      if (typeof entry === "string") {
        p.push(`<si><t>${escapeXml(entry)}</t></si>`);
      } else {
        p.push(`<si>${buildRstXml(entry)}</si>`);
      }
    }

    p.push("</sst>");
    return p.join("");
  },

  parse(el, _ctx) {
    const entries: (string | RichTextOptions)[] = [];

    for (const si of el.elements ?? []) {
      if (si.name !== "si") continue;

      // Simple: <si><t>text</t></si>
      const t = findChild(si, "t");
      if (t) {
        entries.push(textOf(t) ?? "");
        continue;
      }

      // Rich text: <si><r>...</r>...</si>
      const runs: { text: string }[] = [];
      for (const r of si.elements ?? []) {
        if (r.name !== "r") continue;
        const rt = findChild(r, "t");
        if (rt) {
          runs.push({ text: textOf(rt) ?? "" });
        }
      }
      if (runs.length > 0) {
        entries.push({ runs });
      }
    }

    return {
      entries,
      uniqueCount: entries.length,
    } as Record<string, unknown>;
  },
};
