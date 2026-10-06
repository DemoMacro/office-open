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
  SharedStringExtensionOptions,
  RichTextRunProperty,
  RichTextOptions,
  RichTextRunOptions,
  RichTextRunPropertiesOptions,
} from "./worksheet";

/** String or rich text entry in the SST. */
type SstEntry = string | RichTextOptions;

const MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const W14_NS = "http://schemas.microsoft.com/office/word/2008/9/16/wordprocessingDrawing";

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
  const partsByTag: Partial<Record<RichTextRunProperty, string>> = {};
  if (pr.bold) partsByTag.b = "<b/>";
  if (pr.italic) partsByTag.i = "<i/>";
  if (pr.strike) partsByTag.strike = "<strike/>";
  if (pr.outline) partsByTag.outline = "<outline/>";
  if (pr.shadow) partsByTag.shadow = "<shadow/>";
  if (pr.condense) partsByTag.condense = "<condense/>";
  if (pr.extend) partsByTag.extend = "<extend/>";
  // val="none" is explicit: a bare <u/> means underline single, so omitting
  // the attribute would flip none → single on parse.
  if (pr.underline === "single") partsByTag.u = "<u/>";
  else if (pr.underline) partsByTag.u = `<u val="${pr.underline}"/>`;
  if (pr.size !== undefined) partsByTag.sz = `<sz val="${pr.size}"/>`;
  if (pr.color) {
    // parseRPr encodes the non-rgb channels in the same string: a short bare
    // number (≤3 digits) is the legacy palette index, "theme:N" a theme slot.
    // Longer digit strings are hex colors ("008000" is green, not index
    // 8000). They must go back to their own attributes — rgb accepts only 8
    // hex chars (AARRGGBB), and rgb="81" makes Excel refuse the whole package.
    const colorAttrs: string[] = [];
    if (/^\d{1,3}$/.test(pr.color)) {
      colorAttrs.push(`indexed="${Number(pr.color)}"`);
    } else if (pr.color.startsWith("theme:")) {
      colorAttrs.push(`theme="${escapeXml(pr.color.slice(6))}"`);
    } else {
      // ST_UnsignedIntHex requires 8 hex chars (AARRGGBB).
      // Auto-prefix FF (fully opaque) when user provides 6-char RGB.
      const rgb = pr.color.length === 6 ? `FF${pr.color}` : pr.color;
      colorAttrs.push(`rgb="${escapeXml(rgb)}"`);
    }
    if (pr.colorTintRaw !== undefined) colorAttrs.push(`tint="${pr.colorTintRaw}"`);
    else if (pr.colorTint !== undefined) colorAttrs.push(`tint="${pr.colorTint}"`);
    partsByTag.color = `<color ${colorAttrs.join(" ")}/>`;
  }
  if (pr.font) partsByTag.rFont = `<rFont val="${escapeXml(pr.font)}"/>`;
  if (pr.charset !== undefined) partsByTag.charset = `<charset val="${pr.charset}"/>`;
  if (pr.family !== undefined) partsByTag.family = `<family val="${pr.family}"/>`;
  if (pr.vertAlign) partsByTag.vertAlign = `<vertAlign val="${pr.vertAlign}"/>`;
  if (pr.scheme) partsByTag.scheme = `<scheme val="${pr.scheme}"/>`;
  const order = pr.propertyOrder ?? [
    "b",
    "i",
    "strike",
    "outline",
    "shadow",
    "condense",
    "extend",
    "u",
    "sz",
    "color",
    "rFont",
    "charset",
    "family",
    "vertAlign",
    "scheme",
  ];
  const parts = order.flatMap((tag) => {
    const part = partsByTag[tag];
    return part ? [part] : [];
  });
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

function sharedStringExtensionXml(extension: SharedStringExtensionOptions): string {
  let elements = "";
  if (extension.placeholder) {
    const placeholder = extension.placeholder;
    const attrs: string[] = [];
    if (placeholder.processContent !== undefined)
      attrs.push(` mc:ProcessContent="${escapeXml(placeholder.processContent)}"`);
    if (placeholder.preserveAttributes !== undefined)
      attrs.push(` mc:PreserveAttributes="${escapeXml(placeholder.preserveAttributes)}"`);
    const textAttrs: string[] = [];
    if (placeholder.textAttributes?.a !== undefined)
      textAttrs.push(` w14:a="${escapeXml(placeholder.textAttributes.a)}"`);
    if (placeholder.textAttributes?.b !== undefined)
      textAttrs.push(` w14:b="${escapeXml(placeholder.textAttributes.b)}"`);
    if (placeholder.textAttributes?.c !== undefined)
      textAttrs.push(` w14:c="${escapeXml(placeholder.textAttributes.c)}"`);
    elements +=
      `<w14:placeholder${attrs.join("")}><t${textAttrs.join("")}>` +
      `${escapeXml(placeholder.text)}</t></w14:placeholder>`;
  }
  if (extension.no) elements += "<w14:no/>";
  return elements;
}

function siXml(entry: RichTextOptions): string {
  const extension = entry.wordDrawingExtension;
  const attributes = extension
    ? ` xmlns:mc="${MC_NS}" mc:Ignorable="w14" xmlns:w14="${W14_NS}"` +
      (extension.attribute !== undefined ? ` w14:attr="${escapeXml(extension.attribute)}"` : "")
    : "";
  return `<si${attributes}>${sharedStringExtensionXml(extension ?? {})}${buildRstXml(entry)}</si>`;
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
  private sourceCount?: number;
  private sourceUniqueCount?: number;

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

  /** Preserve the source's total-reference count across a round trip. */
  public setSourceCount(value: number): void {
    this.sourceCount = value;
  }

  public setSourceUniqueCount(value: number): void {
    this.sourceUniqueCount = value;
  }

  /** Return a serializable snapshot for the descriptor. */
  public toDescriptorOptions(): {
    entries: SstEntry[];
    count?: number;
    uniqueCount?: number;
  } {
    return {
      entries: this.entries,
      ...(this.sourceCount !== undefined ? { count: this.sourceCount } : {}),
      ...(this.sourceUniqueCount !== undefined ? { uniqueCount: this.sourceUniqueCount } : {}),
    };
  }

  /** Serialize to xl/sharedStrings.xml content (without XML declaration). */
  public serialize(): string {
    return serializeSst(
      this.entries,
      this.sourceCount ?? this.entries.length,
      this.sourceUniqueCount ?? this.entries.length,
    );
  }
}

function serializeSst(
  entries: (string | RichTextOptions)[],
  referenceCount?: number,
  uniqueCount?: number,
): string {
  const p: string[] = [
    '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"',
    referenceCount === undefined ? "" : ` count="${referenceCount}"`,
    uniqueCount === undefined ? "" : ` uniqueCount="${uniqueCount}"`,
  ];
  p.push(">");
  for (const entry of entries) {
    if (typeof entry === "string") {
      p.push(`<si>${tElement(entry)}</si>`);
    } else {
      p.push(siXml(entry));
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
  /** Total string-cell references (<sst/@count>); defaults to entries.length. */
  count?: number;
  /** Unique string entries (<sst/@uniqueCount>); defaults to entries.length. */
  uniqueCount?: number;
}

// ── Descriptor ──

export const sharedStringsDesc: CustomDescriptor<SharedStringsDocOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    if (opts.entries.length === 0) return undefined;
    return serializeSst(opts.entries, opts.count, opts.uniqueCount);
  },

  parse(el, _ctx) {
    const entries: (string | RichTextOptions)[] = [];

    for (const si of el.elements ?? []) {
      if (si.name !== "si") continue;
      const placeholderEl = findChild(si, "w14:placeholder");
      const placeholderText = placeholderEl ? findChild(placeholderEl, "t") : undefined;
      const wordDrawingExtension: SharedStringExtensionOptions | undefined =
        attr(si, "w14:attr") !== undefined || placeholderEl || findChild(si, "w14:no")
          ? {
              ...(attr(si, "w14:attr") !== undefined ? { attribute: attr(si, "w14:attr") } : {}),
              ...(placeholderEl
                ? {
                    placeholder: {
                      ...(attr(placeholderEl, "mc:ProcessContent") !== undefined
                        ? { processContent: attr(placeholderEl, "mc:ProcessContent") }
                        : {}),
                      ...(attr(placeholderEl, "mc:PreserveAttributes") !== undefined
                        ? { preserveAttributes: attr(placeholderEl, "mc:PreserveAttributes") }
                        : {}),
                      text: textOf(placeholderText ?? placeholderEl) ?? "",
                      ...(placeholderText?.attributes
                        ? {
                            textAttributes: {
                              a: attr(placeholderText, "w14:a"),
                              b: attr(placeholderText, "w14:b"),
                              c: attr(placeholderText, "w14:c"),
                            },
                          }
                        : {}),
                    },
                  }
                : {}),
              ...(findChild(si, "w14:no") ? { no: true } : {}),
            }
          : undefined;

      // Simple: <si><t>text</t></si> — phonetic children may still trail
      // (CT_Rst allows t + rPh* + phoneticPr without any r runs), in which
      // case the entry stays a RichTextOptions to carry them.
      const t = findChild(si, "t");
      const hasPhonetic = (si.elements ?? []).some(
        (e) => e.name === "rPh" || e.name === "phoneticPr",
      );
      if (t && !hasPhonetic && !wordDrawingExtension) {
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
        if (wordDrawingExtension) entry.wordDrawingExtension = wordDrawingExtension;
        entries.push(entry);
      } else if (t) {
        // Plain text with trailing phonetics — text + rPh*/phoneticPr.
        const entry: RichTextOptions = { text: textOf(t) ?? "" };
        if (phonetics.length > 0) entry.phonetics = phonetics;
        if (phoneticProperties) entry.phoneticProperties = phoneticProperties;
        if (wordDrawingExtension) entry.wordDrawingExtension = wordDrawingExtension;
        entries.push(entry);
      } else if (!hasPhonetic) {
        entries.push({});
      }
    }

    return { entries };
  },
};

/** Parse CT_RPrElt (run properties inside shared strings r element). */
export function parseRPr(el: XmlElement): RichTextRunPropertiesOptions {
  const result: RichTextRunPropertiesOptions = {};
  const propertyOrder: RichTextRunProperty[] = [];
  for (const child of el.elements ?? []) {
    switch (child.name) {
      case "rFont":
      case "charset":
      case "family":
      case "b":
      case "i":
      case "strike":
      case "outline":
      case "shadow":
      case "condense":
      case "extend":
      case "color":
      case "sz":
      case "u":
      case "vertAlign":
      case "scheme":
        propertyOrder.push(child.name);
        break;
    }
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
        const tint = attr(child, "tint");
        if (tint !== undefined && Number.isFinite(Number(tint))) result.colorTint = Number(tint);
        if (tint !== undefined) result.colorTintRaw = tint;
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
  if (propertyOrder.length > 0) result.propertyOrder = propertyOrder;
  return result;
}
