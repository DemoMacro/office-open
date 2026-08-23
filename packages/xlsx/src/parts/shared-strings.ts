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
 * Serialize a CT_Rst text element. Excel requires xml:space="preserve" on a
 * `<t>` whose text has leading or trailing whitespace — without it the whole
 * part is rejected on open. The attribute is derived from the text itself, so
 * parse needs no field for it.
 */
export function tElement(text: string): string {
  const open = /^\s|\s$/.test(text) ? '<t xml:space="preserve">' : "<t>";
  return `${open}${escapeXml(text)}</t>`;
}

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
    // parseRPr encodes the non-rgb channels in the same string: a short bare
    // number (≤3 digits) is the legacy palette index, "theme:N" a theme slot.
    // Longer digit strings are hex colors ("008000" is green, not index
    // 8000). They must go back to their own attributes — rgb accepts only 8
    // hex chars (AARRGGBB), and rgb="81" makes Excel refuse the whole package.
    if (/^\d{1,3}$/.test(pr.color)) {
      parts.push(`<color indexed="${Number(pr.color)}"/>`);
    } else if (pr.color.startsWith("theme:")) {
      parts.push(`<color theme="${escapeXml(pr.color.slice(6))}"/>`);
    } else {
      // ST_UnsignedIntHex requires 8 hex chars (AARRGGBB).
      // Auto-prefix FF (fully opaque) when user provides 6-char RGB.
      const rgb = pr.color.length === 6 ? `FF${pr.color}` : pr.color;
      parts.push(`<color rgb="${escapeXml(rgb)}"/>`);
    }
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
      parts.push(`<r>${rPr}${tElement(run.text)}</r>`);
    }
  } else if (rst.text !== undefined) {
    parts.push(tElement(rst.text));
  }
  // rPh (phonetics)
  if (rst.phonetics) {
    for (const ph of rst.phonetics) {
      parts.push(`<rPh sb="${ph.startByte}" eb="${ph.endByte}">${tElement(ph.text)}</rPh>`);
    }
  }
  if (rst.phoneticProperties) {
    const pp = rst.phoneticProperties;
    const attrs: string[] = [`fontId="${pp.fontId}"`];
    if (pp.type) attrs.push(`type="${pp.type}"`);
    if (pp.alignment) attrs.push(`alignment="${pp.alignment}"`);
    parts.push(`<phoneticPr ${attrs.join(" ")}/>`);
  }
  return parts.join("");
}

export class SharedStrings {
  private entries: SstEntry[] = [];
  /** Dedup map for plain strings only. */
  private indexMap = new Map<string, number>();
  /**
   * Identity dedup map for rich text — round-tripped cells hold the same
   * entry object the loaded table already contains, so registering them must
   * resolve back to the original index instead of appending a duplicate si.
   */
  private richIndexMap = new Map<RichTextOptions, number>();

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
   * Register a rich text entry and return its index. The same object
   * (identity) resolves to its existing index; distinct objects are appended.
   */
  public registerRich(rst: RichTextOptions): number {
    const existing = this.richIndexMap.get(rst);
    if (existing !== undefined) return existing;

    const idx = this.entries.length;
    this.entries.push(rst);
    this.richIndexMap.set(rst, idx);
    return idx;
  }

  /**
   * Bulk-load parsed template entries, preserving their original indices so
   * existing cell references stay valid. Plain strings populate the dedup map
   * (first occurrence wins); rich-text entries populate the identity map.
   *
   * Used by patch and by generate() for a parsed workbook, so cells keep
   * pointing at the source table instead of re-registering flattened text.
   */
  public loadEntries(entries: SharedStringsDocOptions["entries"]): void {
    for (const entry of entries) {
      const idx = this.entries.length;
      this.entries.push(entry);
      if (typeof entry === "string") {
        if (!this.indexMap.has(entry)) this.indexMap.set(entry, idx);
      } else if (!this.richIndexMap.has(entry)) {
        this.richIndexMap.set(entry, idx);
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
      p.push(`<si>${tElement(entry)}</si>`);
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

      // Simple: <si><t>text</t></si> — phonetic children may still trail
      // (CT_Rst allows t + rPh* + phoneticPr without any r runs), in which
      // case the entry stays a RichTextOptions to carry them.
      const t = findChild(si, "t");
      const hasPhonetic = (si.elements ?? []).some(
        (e) => e.name === "rPh" || e.name === "phoneticPr",
      );
      if (t && !hasPhonetic) {
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

      // Phonetics: <rPh sb="..." eb="..."><t>...</t></rPh> + trailing phoneticPr
      const phonetics: { startByte: number; endByte: number; text: string }[] = [];
      let phoneticProperties: RichTextOptions["phoneticProperties"];
      for (const rPh of si.elements ?? []) {
        if (rPh.name === "rPh") {
          const sb = attrNum(rPh, "sb") ?? 0;
          const eb = attrNum(rPh, "eb") ?? 0;
          const rPhT = findChild(rPh, "t");
          phonetics.push({ startByte: sb, endByte: eb, text: rPhT ? (textOf(rPhT) ?? "") : "" });
        } else if (rPh.name === "phoneticPr") {
          const fontId = attrNum(rPh, "fontId");
          if (fontId !== undefined) {
            const pp: NonNullable<RichTextOptions["phoneticProperties"]> = { fontId };
            const type = attr(rPh, "type");
            if (type) pp.type = type as NonNullable<RichTextOptions["phoneticProperties"]>["type"];
            const align = attr(rPh, "alignment");
            if (align)
              pp.alignment = align as NonNullable<
                RichTextOptions["phoneticProperties"]
              >["alignment"];
            phoneticProperties = pp;
          }
        }
      }

      if (runs.length > 0) {
        const entry: RichTextOptions = { runs };
        if (phonetics.length > 0) entry.phonetics = phonetics;
        if (phoneticProperties) entry.phoneticProperties = phoneticProperties;
        entries.push(entry);
      } else if (t) {
        // Plain text with trailing phonetics — text + rPh*/phoneticPr.
        const entry: RichTextOptions = { text: textOf(t) ?? "" };
        if (phonetics.length > 0) entry.phonetics = phonetics;
        if (phoneticProperties) entry.phoneticProperties = phoneticProperties;
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
