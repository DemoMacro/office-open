/**
 * Shared Strings Table — generates xl/sharedStrings.xml.
 *
 * XLSX stores repeated string values in a central table to reduce file size.
 * Cells reference strings by index into this table.
 *
 * @module
 */
import { BaseXmlComponent } from "@file/xml-components";
import type { Context } from "@file/xml-components";
import { escapeXml } from "@office-open/xml";

import type { RichTextOptions } from "./worksheet";

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
  if (pr.underline) {
    if (pr.underline === "none") {
      parts.push("<u/>");
    } else {
      parts.push(`<u val="${pr.underline}"/>`);
    }
  }
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
      parts.push(`<rPh sb="${ph.sb}" eb="${ph.eb}"><t>${escapeXml(ph.text)}</t></rPh>`);
    }
  }
  return parts.join("");
}

export class SharedStrings extends BaseXmlComponent {
  private readonly entries: SstEntry[] = [];
  /** Dedup map for plain strings only. Rich text is not deduped. */
  private readonly indexMap = new Map<string, number>();

  public constructor() {
    super("sst");
  }

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

  public get count(): number {
    return this.entries.length;
  }

  /**
   * Zero-allocation fast path: directly concatenate XML string.
   * Bypasses the IXmlableObject intermediate tree entirely.
   */
  public override toXml(_context: Context): string {
    const p: string[] = [
      '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"',
      ` count="${this.entries.length}" uniqueCount="${this.indexMap.size}">`,
    ];
    for (const entry of this.entries) {
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
}
