/**
 * Shared Strings Table — generates xl/sharedStrings.xml.
 *
 * XLSX stores repeated string values in a central table to reduce file size.
 * Cells reference strings by index into this table.
 *
 * @module
 */
import { parseOnOff } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { escapeXml, findChild, attr, attrNum, textOf } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";

import { parseColorHex } from "./styles/parse";
import type {
  RichTextOptions,
  RichTextRunOptions,
  RichTextRunPropertiesOptions,
} from "./worksheet";

/** String or rich text entry in the SST. */
type SstEntry = string | RichTextOptions;

/**
 * Build rich text run properties XML (CT_RPrElt).
 * Exported for reuse by Comments and other components.
 */
export function buildRPrXml(
  pr: NonNullable<RichTextOptions["runs"]>[number]["properties"],
): string {
  if (!pr) return "";
  const parts: string[] = [];
  if (pr.font) parts.push(`<rFont val="${escapeXml(pr.font)}"/>`);
  if (pr.charset !== undefined) parts.push(`<charset val="${pr.charset}"/>`);
  if (pr.family !== undefined) parts.push(`<family val="${pr.family}"/>`);
  if (pr.bold) parts.push("<b/>");
  if (pr.italic) parts.push("<i/>");
  if (pr.strike) parts.push("<strike/>");
  if (pr.outline) parts.push("<outline/>");
  if (pr.shadow) parts.push("<shadow/>");
  if (pr.condense) parts.push("<condense/>");
  if (pr.extend) parts.push("<extend/>");
  if (pr.color) {
    // ST_UnsignedIntHex requires 8 hex chars (AARRGGBB).
    // Auto-prefix FF (fully opaque) when user provides 6-char RGB.
    const rgb = pr.color.length === 6 ? `FF${pr.color}` : pr.color;
    parts.push(`<color rgb="${escapeXml(rgb)}"/>`);
  }
  if (pr.size !== undefined) parts.push(`<sz val="${pr.size}"/>`);
  // val="none" is explicit: a bare <u/> means underline single, so omitting
  // the attribute would flip none → single on parse.
  if (pr.underline) parts.push(`<u val="${pr.underline}"/>`);
  if (pr.vertAlign) parts.push(`<vertAlign val="${pr.vertAlign}"/>`);
  if (pr.scheme) parts.push(`<scheme val="${pr.scheme}"/>`);
  return parts.length > 0 ? `<rPr>${parts.join("")}</rPr>` : "";
}

/** Build a CT_Rst XML string from RichTextOptions. */
export function buildRstXml(rst: RichTextOptions): string {
  const parts: string[] = [];
  if (rst.runs && rst.runs.length > 0) {
    for (const run of rst.runs) {
      const rPr = buildRPrXml(run.properties);
      parts.push(`<r>${rPr}<t>${escapeXml(run.text)}</t></r>`);
    }
  } else if (rst.text !== undefined) {
    parts.push(`<t>${escapeXml(rst.text)}</t>`);
  }
  // rPh (phonetics)
  if (rst.phonetics) {
    for (const ph of rst.phonetics) {
      parts.push(`<rPh sb="${ph.startByte}" eb="${ph.endByte}"><t>${escapeXml(ph.text)}</t></rPh>`);
    }
  }
  return parts.join("");
}

export class SharedStrings {
  private entries: SstEntry[] = [];
  /** Dedup map for plain strings only. Rich text is not deduped. */
  private indexMap = new Map<string, number>();

  /**
   * Register a plain string and return its index.
   * Returns existing index if the string is already registered.
   */
  public register(s: string): number {
    const existing = this.indexMap.get(s);
    if (existing !== undefined) return existing;

    const idx = this.entries.length;
    this.entries.push(s);
    this.indexMap.set(s, idx);
    return idx;
  }

  /**
   * Register a rich text entry and return its index.
   * Rich text is not deduped (each call creates a new entry).
   */
  public registerRich(rst: RichTextOptions): number {
    const idx = this.entries.length;
    this.entries.push(rst);
    return idx;
  }

  /**
   * Bulk-load parsed template entries, preserving their original indices so
   * existing cell references stay valid. Plain strings populate the dedup map
   * (first occurrence wins); rich-text entries are appended unchanged.
   *
   * Used by patch to extend an existing workbook's shared strings, so appended
   * worksheets continue registering strings at the correct offset.
   */
  public loadEntries(entries: SharedStringsDocOptions["entries"]): void {
    for (const entry of entries) {
      const idx = this.entries.length;
      this.entries.push(entry);
      if (typeof entry === "string" && !this.indexMap.has(entry)) {
        this.indexMap.set(entry, idx);
      }
    }
  }

  public get count(): number {
    return this.entries.length;
  }

  /** Return a serializable snapshot for the descriptor. */
  public toDescriptorOptions(): { entries: SstEntry[] } {
    return { entries: this.entries };
  }

  /** Serialize to xl/sharedStrings.xml content (without XML declaration). */
  public serialize(): string {
    return serializeSstEntries(this.entries);
  }
}

/** Serialize entry list to the <sst> part body shared by both emit paths. */
function serializeSstEntries(entries: (string | RichTextOptions)[]): string {
  const p: string[] = [
    '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"',
    ` count="${entries.length}" uniqueCount="${entries.length}">`,
  ];
  for (const entry of entries) {
    if (typeof entry === "string") {
      p.push(`<si><t>${escapeXml(entry)}</t></si>`);
    } else {
      // Rich text (CT_Rst)
      p.push(`<si>${buildRstXml(entry)}</si>`);
    }
  }
  p.push("</sst>");
  return p.join("");
}

// ── Descriptor Types ──

/** Serializable snapshot of the shared string table. uniqueCount is not
 * stored — it always equals entries.length and is derived at emit time. */
export interface SharedStringsDocOptions {
  /** All entries (plain strings and rich text), in registration order. */
  entries: (string | RichTextOptions)[];
}

// ── Descriptor ──

export const sharedStringsDesc: CustomDescriptor<SharedStringsDocOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    if (opts.entries.length === 0) return undefined;
    return serializeSstEntries(opts.entries);
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
      const runs: RichTextRunOptions[] = [];
      for (const r of si.elements ?? []) {
        if (r.name !== "r") continue;
        const rt = findChild(r, "t");
        if (rt) {
          const rPrEl = findChild(r, "rPr");
          const run: RichTextRunOptions = { text: textOf(rt) ?? "" };
          if (rPrEl) run.properties = parseRPr(rPrEl);
          runs.push(run);
        }
      }

      // Phonetics: <rPh sb="..." eb="..."><t>...</t></rPh>
      const phonetics: { startByte: number; endByte: number; text: string }[] = [];
      for (const rPh of si.elements ?? []) {
        if (rPh.name !== "rPh") continue;
        const sb = attrNum(rPh, "sb") ?? 0;
        const eb = attrNum(rPh, "eb") ?? 0;
        const rPhT = findChild(rPh, "t");
        phonetics.push({ startByte: sb, endByte: eb, text: rPhT ? (textOf(rPhT) ?? "") : "" });
      }

      if (runs.length > 0) {
        const entry: RichTextOptions = { runs };
        if (phonetics.length > 0) entry.phonetics = phonetics;
        entries.push(entry);
      }
    }

    return { entries };
  },
};

/** Parse CT_RPrElt (run properties inside shared strings r element). */
export function parseRPr(el: XmlElement): RichTextRunPropertiesOptions {
  const result: RichTextRunPropertiesOptions = {};
  for (const child of el.elements ?? []) {
    switch (child.name) {
      case "rFont":
        result.font = attr(child, "val") ?? undefined;
        break;
      case "charset":
        result.charset = attrNum(child, "val");
        break;
      case "family":
        result.family = attrNum(child, "val");
        break;
      case "b":
        result.bold = parseOnOff(attr(child, "val")) ?? true;
        break;
      case "i":
        result.italic = parseOnOff(attr(child, "val")) ?? true;
        break;
      case "strike":
        result.strike = true;
        break;
      case "outline":
        result.outline = true;
        break;
      case "shadow":
        result.shadow = true;
        break;
      case "condense":
        result.condense = true;
        break;
      case "extend":
        result.extend = true;
        break;
      case "color": {
        const rgb = parseColorHex(child);
        if (rgb) {
          result.color = rgb;
        } else {
          const indexed = attrNum(child, "indexed");
          if (indexed !== undefined) result.color = String(indexed);
          else {
            const theme = attr(child, "theme");
            if (theme !== undefined) result.color = `theme:${theme}`;
          }
        }
        break;
      }
      case "sz":
        result.size = attrNum(child, "val");
        break;
      case "u": {
        const uVal = attr(child, "val");
        result.underline =
          (uVal as RichTextRunPropertiesOptions["underline"] | undefined) ?? "single";
        break;
      }
      case "vertAlign":
        result.vertAlign = attr(child, "val") as
          | RichTextRunPropertiesOptions["vertAlign"]
          | undefined;
        break;
      case "scheme":
        result.scheme = attr(child, "val") as RichTextRunPropertiesOptions["scheme"] | undefined;
        break;
    }
  }
  return result;
}
