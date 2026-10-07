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

import { colorAttributes, parseColorOptions } from "../shared/color";
import type {
  SharedStringExtensionOptions,
  RichTextRunProperty,
  RichTextOptions,
  RichTextRunOptions,
  RichTextRunPropertiesOptions,
  PhoneticRunOptions,
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
 * Serialize a `<t>` with an explicit source xml:space form. `preserve` keeps
 * the attribute, `default` drops it even when the text has outer whitespace
 * (matching a source that omitted the attribute), and an undefined form falls
 * back to the derived fresh-authoring behavior of {@link tElement}.
 */
export function tElementRaw(text: string, spaceRaw: string | undefined): string {
  if (spaceRaw === "default") return `<t>${escapeXml(text)}</t>`;
  if (spaceRaw === "preserve" || /^\s|\s$/.test(text))
    return `<t xml:space="preserve">${escapeXml(text)}</t>`;
  return `<t>${escapeXml(text)}</t>`;
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
  if (pr.bold)
    partsByTag.b = pr.boldValRaw !== undefined ? `<b val="${escapeXml(pr.boldValRaw)}"/>` : "<b/>";
  if (pr.italic) partsByTag.i = "<i/>";
  if (pr.strike) partsByTag.strike = "<strike/>";
  if (pr.outline) partsByTag.outline = "<outline/>";
  if (pr.shadow) partsByTag.shadow = "<shadow/>";
  if (pr.condense) partsByTag.condense = "<condense/>";
  if (pr.extend) partsByTag.extend = "<extend/>";
  // val="none" is explicit: a bare <u/> means underline single, so omitting
  // the attribute would flip none → single on parse.
  if (pr.underlineValRaw !== undefined)
    partsByTag.u = `<u val="${escapeXml(pr.underlineValRaw)}"/>`;
  else if (pr.underline === "single") partsByTag.u = "<u/>";
  else if (pr.underline) partsByTag.u = `<u val="${pr.underline}"/>`;
  if (pr.size !== undefined)
    partsByTag.sz =
      pr.sizeValRaw !== undefined
        ? `<sz val="${escapeXml(pr.sizeValRaw)}"/>`
        : `<sz val="${pr.size}"/>`;
  if (pr.color) {
    partsByTag.color = `<color ${colorAttributes(pr.color)}/>`;
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
  // CT_Rst child order: t, r*, rPh*, phoneticPr. A source-faithful entry may
  // carry plain text and runs together.
  if (rst.text !== undefined) parts.push(tElementRaw(rst.text, rst.textSpaceRaw));
  if (rst.runs && rst.runs.length > 0) {
    for (const run of rst.runs) {
      const rPr = buildRPrXml(run.properties);
      parts.push(`<r>${rPr}${tElementRaw(run.text, run.textSpaceRaw)}</r>`);
    }
  }
  // rPh (phonetics)
  if (rst.phonetics) {
    for (const ph of rst.phonetics) {
      const inner =
        ph.runs && ph.runs.length > 0
          ? ph.runs
              .map((run) => `<r>${buildRPrXml(run.properties)}${tElement(run.text)}</r>`)
              .join("")
          : tElement(ph.text);
      parts.push(`<rPh sb="${ph.startByte}" eb="${ph.endByte}">${inner}</rPh>`);
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
  /** Whether parsed source entries were loaded (round-trip mode). */
  public sourceLoaded = false;

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
    this.sourceLoaded = true;
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
    return serializeSst(opts.entries, opts.count, opts.uniqueCount);
  },

  parse(el, _ctx) {
    const entries: (string | RichTextOptions)[] = [];

    for (const si of el.elements ?? []) {
      if (si.name !== "si" && si.name !== "sstItem") continue;
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

      const t = findChild(si, "t");
      const runs: RichTextRunOptions[] = [];
      for (const r of si.elements ?? []) {
        if (r.name !== "r") continue;
        const rt = findChild(r, "t");
        if (rt) {
          const rPrEl = findChild(r, "rPr");
          const run: RichTextRunOptions = { text: textOf(rt) ?? "" };
          if (attr(rt, "xml:space") !== undefined) run.textSpaceRaw = attr(rt, "xml:space");
          if (rPrEl) run.properties = parseRPr(rPrEl);
          runs.push(run);
        }
      }

      // Phonetics: <rPh sb="..." eb="...">…</rPh> + trailing phoneticPr. The
      // XSD form wraps one <t>; Excel-compatible producers may wrap CT_RElt
      // <r> runs instead, which parse into the phonetic `runs` field.
      const phonetics: PhoneticRunOptions[] = [];
      let phoneticProperties: RichTextOptions["phoneticProperties"];
      for (const rPh of si.elements ?? []) {
        if (rPh.name === "rPh") {
          const sb = attrNum(rPh, "sb") ?? 0;
          const eb = attrNum(rPh, "eb") ?? 0;
          const rPhT = findChild(rPh, "t");
          const rPhRuns: RichTextRunOptions[] = [];
          for (const child of rPh.elements ?? []) {
            if (child.name !== "r") continue;
            const rt = findChild(child, "t");
            if (!rt) continue;
            const run: RichTextRunOptions = { text: textOf(rt) ?? "" };
            const rPrEl = findChild(child, "rPr");
            if (rPrEl) run.properties = parseRPr(rPrEl);
            rPhRuns.push(run);
          }
          const phonetic: PhoneticRunOptions = {
            startByte: sb,
            endByte: eb,
            text: rPhT ? (textOf(rPhT) ?? "") : "",
          };
          if (rPhRuns.length > 0) phonetic.runs = rPhRuns;
          phonetics.push(phonetic);
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

      if (
        t ||
        runs.length > 0 ||
        phonetics.length > 0 ||
        phoneticProperties !== undefined ||
        wordDrawingExtension !== undefined
      ) {
        const entry: RichTextOptions = {};
        if (t) {
          const text = textOf(t) ?? "";
          entry.text = text;
          const sourceSpace = t.attributes?.["xml:space"];
          if (sourceSpace !== undefined) {
            entry.textSpaceRaw = sourceSpace === "preserve" ? "preserve" : "default";
          } else if (/^\s|\s$/.test(text)) {
            entry.textSpaceRaw = "default";
          }
        }
        if (runs.length > 0) entry.runs = runs;
        if (phonetics.length > 0) entry.phonetics = phonetics;
        if (phoneticProperties) entry.phoneticProperties = phoneticProperties;
        if (wordDrawingExtension) entry.wordDrawingExtension = wordDrawingExtension;
        // Fast path: unflagged plain text stays a plain string so cell values
        // and SST value dedup keep their previous shape.
        if (
          entry.runs === undefined &&
          entry.phonetics === undefined &&
          entry.phoneticProperties === undefined &&
          entry.wordDrawingExtension === undefined &&
          entry.textSpaceRaw === undefined
        ) {
          entries.push(entry.text ?? "");
        } else {
          entries.push(entry);
        }
      } else {
        // Empty si — keep the slot so subsequent SST indices stay aligned.
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
        if (attr(child, "val") !== undefined) result.boldValRaw = attr(child, "val");
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
        result.color = parseColorOptions(child);
        break;
      }
      case "sz":
        result.size = attrNum(child, "val");
        if (attr(child, "val") !== undefined) result.sizeValRaw = attr(child, "val");
        break;
      case "u": {
        const uVal = attr(child, "val");
        if (uVal !== undefined) result.underlineValRaw = uVal;
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
